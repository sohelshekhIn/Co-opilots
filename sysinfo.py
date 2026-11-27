"""Read-only facts about this PC for Copilot. Commands are fixed here; Copilot only picks a topic name."""

import ctypes
import json
import os
import re
import subprocess
import time
from pathlib import Path

import pygetwindow

_LIMIT = 3000
# ponytail: one JSON line per task, read back whole; fine for thousands of tasks. Upgrade path is trimming old lines.
HISTORY = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "CoPilotComputerUse" / "history.jsonl"

_NOW = r"""
"Time: $((Get-Date).ToString('dddd yyyy-MM-dd HH:mm'))"
try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $o = $null }
if ($o) {
  $items = $o.GetNamespace('MAPI').GetDefaultFolder(9).Items; $items.IncludeRecurrences = $true; $items.Sort('[Start]')
  $now = Get-Date; $current = @(); $next = $null
  $until = $now.Date.AddDays(14)
  foreach ($m in $items.Restrict("[End] > '$($now.ToString('g'))' AND [Start] < '$($until.ToString('g'))'")) {
    if ($m.AllDayEvent) { continue }
    if ($m.Start -le $now) { $current += $m } elseif (-not $next) { $next = $m; break } }
  if ($current) { $current | ForEach-Object { "Meeting now: {0} ({1:HH:mm}-{2:HH:mm})" -f $_.Subject, $_.Start, $_.End } } else { "Meeting now: none" }
  if ($next) { "Next meeting: {0} at {1:ddd HH:mm} (in {2:N0} min)" -f $next.Subject, $next.Start, ($next.Start - $now).TotalMinutes } else { "Next meeting: none in the next 14 days" }
} else { "Calendar: classic Outlook is not running" }
"Recently opened files:"
Get-ChildItem "$env:APPDATA\Microsoft\Windows\Recent" -Filter *.lnk -ErrorAction SilentlyContinue |
  Where-Object { $_.BaseName -notlike 'ms-*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 10 |
  ForEach-Object { "  {0:MM-dd HH:mm} {1}" -f $_.LastWriteTime, $_.BaseName }
"""


def record_history(task, answer, seconds):
    try:
        HISTORY.parent.mkdir(parents=True, exist_ok=True)
        entry = {"at": time.strftime("%Y-%m-%d %H:%M"), "task": task[:500], "answer": str(answer)[:800],
                 "seconds": round(seconds)}
        with HISTORY.open("a", encoding="utf-8") as file:
            file.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except OSError:
        pass


def recent_history(count):
    try:
        lines = HISTORY.read_text(encoding="utf-8").splitlines()[-count:]
    except OSError:
        return ""
    out = []
    for line in lines:
        try:
            entry = json.loads(line)
        except ValueError:
            continue
        answer = " ".join(entry.get("answer", "").split())
        out.append(f"- {entry.get('at')}: {entry.get('task')} -> {answer[:300]}")
    return "\n".join(out)


def _front_windows(count=3):
    """Top visible windows in z-order, skipping this app's own windows and its hidden browser."""
    import desktop
    user32, me = ctypes.windll.user32, os.getpid()
    screens, found = desktop.monitors(), []
    for window in pygetwindow.getAllWindows():
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(window._hWnd, ctypes.byref(pid))
        title = window.title.strip()
        if (not title or not window.visible or window.isMinimized or pid.value == me
                or "Chrome for Testing" in title
                or title in ("Copilot Assistant", "Co-opilot", "Developer log", "Program Manager", "Windows Input Experience")):
            continue
        number = desktop.window_monitor(window, screens)
        if number:
            found.append(f"  {title}  [monitor {number}]")
        if len(found) >= count:
            break
    return found

_COMMANDS = {
    "system": r"""
$os = Get-CimInstance Win32_OperatingSystem; $cs = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$up = (Get-Date) - $os.LastBootUpTime
"Computer: $env:COMPUTERNAME  User: $env:USERDOMAIN\$env:USERNAME"
"OS: $($os.Caption) $($os.Version) build $($os.BuildNumber)"
"Model: $($cs.Manufacturer) $($cs.Model)"
"CPU: $($cpu.Name)  load $($cpu.LoadPercentage)%"
"RAM: {0:N1} GB free of {1:N1} GB" -f ($os.FreePhysicalMemory/1MB), ($os.TotalVisibleMemorySize/1MB)
"Uptime: {0}d {1}h {2}m" -f $up.Days, $up.Hours, $up.Minutes
""",
    "disks": r"""
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
  "{0} {1:N1} GB free of {2:N1} GB ({3:N0}% free)" -f $_.DeviceID, ($_.FreeSpace/1GB), ($_.Size/1GB), (100*$_.FreeSpace/[math]::Max($_.Size,1)) }
""",
    "processes": r"""
$cores = [Environment]::ProcessorCount
"CPU right now (% of whole PC):"
Get-CimInstance Win32_PerfFormattedData_PerfProc_Process |
  Where-Object { $_.Name -notin '_Total','Idle' -and $_.PercentProcessorTime -gt 0 } |
  Sort-Object PercentProcessorTime -Descending | Select-Object -First 8 |
  ForEach-Object { "{0,-28} {1,5:N1} %" -f $_.Name, ($_.PercentProcessorTime / $cores) }
"Top by memory:"
Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 12 |
  ForEach-Object { "{0,-28} {1,8:N0} MB  cpu {2,8:N1}s" -f $_.ProcessName, ($_.WorkingSet64/1MB), $_.CPU }
"Top by CPU time:"
Get-Process | Sort-Object CPU -Descending | Select-Object -First 8 |
  ForEach-Object { "{0,-28} cpu {1,8:N1}s" -f $_.ProcessName, $_.CPU }
""",
    "network": r"""
Get-NetIPConfiguration | Where-Object { $_.IPv4Address } | ForEach-Object {
  "{0}: {1} gateway {2} dns {3}" -f $_.InterfaceAlias, ($_.IPv4Address.IPAddress -join ','),
    ($_.IPv4DefaultGateway.NextHop -join ','), ($_.DNSServer.ServerAddresses -join ',') }
$null = ping.exe -n 1 -w 1500 8.8.8.8
"Internet ping 8.8.8.8: $($LASTEXITCODE -eq 0)"
try { $null = Resolve-DnsName login.microsoftonline.com -ErrorAction Stop; "DNS login.microsoftonline.com: ok" } catch { "DNS login.microsoftonline.com: failed" }
""",
    "services": r"""
"Automatic services not running:"
Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" |
  Select-Object -First 25 | ForEach-Object { "{0} ({1}) {2}" -f $_.Name, $_.DisplayName, $_.State }
""",
    "events": r"""
$since = (Get-Date).AddHours(-24)
foreach ($log in 'System','Application') {
  "== $log errors, last 24h =="
  Get-WinEvent -FilterHashtable @{LogName=$log; Level=1,2; StartTime=$since} -MaxEvents 10 -ErrorAction SilentlyContinue |
    ForEach-Object { "{0:MM-dd HH:mm} {1}: {2}" -f $_.TimeCreated, $_.ProviderName, (($_.Message -split "`n")[0]) }
}
""",
    "apps": r"""
$keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object DisplayName |
  Sort-Object DisplayName -Unique | ForEach-Object { "$($_.DisplayName) $($_.DisplayVersion)" }
""",
    "startup": r"""
"Startup apps:"
Get-CimInstance Win32_StartupCommand | ForEach-Object { "{0}  [{1}]" -f $_.Name, $_.Location }
""",
    "battery": r"""
$b = Get-CimInstance Win32_Battery
if (-not $b) { "No battery (desktop or docked without battery info)" } else {
  $b | ForEach-Object { "Battery: $($_.EstimatedChargeRemaining)% status $($_.BatteryStatus) (1=discharging, 2=on AC)" } }
$plan = (powercfg /getactivescheme) -replace '.*\((.*)\).*','$1'
"Power plan: $plan"
""",
    "updates": r"""
"Last installed updates:"
Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10 |
  ForEach-Object { "{0} {1:yyyy-MM-dd} {2}" -f $_.HotFixID, $_.InstalledOn, $_.Description }
$pending = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
"Restart pending for updates: $pending"
""",
    "audio": r"""
"Sound hardware:"
Get-CimInstance Win32_SoundDevice | ForEach-Object { "  {0}  status {1}" -f $_.Name, $_.Status }
"Speakers and microphones (endpoints):"
Get-PnpDevice -Class AudioEndpoint -PresentOnly -ErrorAction SilentlyContinue |
  ForEach-Object { "  {0}  {1}" -f $_.FriendlyName, $_.Status }
$svc = Get-Service Audiosrv, AudioEndpointBuilder -ErrorAction SilentlyContinue
$svc | ForEach-Object { "Service $($_.Name): $($_.Status)" }
""",
    "cameras": r"""
"Cameras:"
Get-PnpDevice -Class Camera, Image -PresentOnly -ErrorAction SilentlyContinue |
  ForEach-Object { "  {0}  {1}" -f $_.FriendlyName, $_.Status }
foreach ($cap in 'webcam', 'microphone') {
  $user = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$cap"
  $machine = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$cap"
  "$cap privacy: device $((Get-ItemProperty $machine -ErrorAction SilentlyContinue).Value), user $((Get-ItemProperty $user -ErrorAction SilentlyContinue).Value)"
  "Apps using the $cap right now:"
  Get-ChildItem $user -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
    if ($p.LastUsedTimeStart -and $p.LastUsedTimeStop -eq 0) { "  " + ($_.PSChildName -replace '#', '\') } }
}
""",
    "outlook": r"""
try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $o = $null }
if (-not $o) { "Classic Outlook is not running here (new Outlook has no local API). Read Outlook on screen instead."; return }
$ns = $o.GetNamespace('MAPI'); $inbox = $ns.GetDefaultFolder(6)
"Unread in Inbox: $($inbox.UnReadItemCount)"
"Latest unread:"
$all = $inbox.Items; $all.Sort('[ReceivedTime]', $true); $unread = $all.Restrict('[UnRead] = True')
for ($i = 1; $i -le [math]::Min(8, $unread.Count); $i++) {
  $m = $unread.Item($i); "  {0:MM-dd HH:mm} {1}: {2}" -f $m.ReceivedTime, $m.SenderName, $m.Subject }
"Today's meetings:"
$items = $ns.GetDefaultFolder(9).Items; $items.IncludeRecurrences = $true; $items.Sort('[Start]')
$day = (Get-Date).Date
$items.Restrict("[Start] >= '$($day.ToString('g'))' AND [Start] < '$($day.AddDays(1).ToString('g'))'") |
  ForEach-Object { "  {0:HH:mm}-{1:HH:mm} {2} ({3})" -f $_.Start, $_.End, $_.Subject, $_.Location }
""",
    "teams": r"""
$p = @(Get-Process ms-teams, Teams -ErrorAction SilentlyContinue)
if (-not $p) { "Teams is not running" } else {
  "Teams running: {0} processes, {1:N0} MB RAM, {2:N0}s CPU" -f $p.Count, (($p | Measure-Object WorkingSet64 -Sum).Sum / 1MB), (($p | Measure-Object CPU -Sum).Sum) }
$pkg = Get-AppxPackage MSTeams -ErrorAction SilentlyContinue
if ($pkg) { "New Teams version $($pkg.Version)" }
$cache = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
if (Test-Path $cache) { "Cache size: {0:N0} MB" -f ((Get-ChildItem $cache -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1MB) }
"Use inspect cameras to see whether Teams is using the camera or microphone right now."
""",
    "printers": r"""
Get-CimInstance Win32_Printer | ForEach-Object {
  "{0}{1}  status {2}  port {3}" -f $_.Name, $(if ($_.Default) {' (default)'} else {''}), $_.PrinterStatus, $_.PortName }
""",
}


def collect(topic):
    if topic == "windows":
        import desktop
        screens = desktop.monitors()
        def where(w):
            if w.isMinimized:
                return "minimized"
            number = desktop.window_monitor(w, screens)
            return f"monitor {number}" if number else "off-screen"
        # The app's own hidden Copilot browser is Chrome for Testing; leave it out.
        lines = sorted({f"{w.title.strip()}  [{where(w)}]" for w in pygetwindow.getAllWindows()
                        if w.title.strip() and w.visible and "Chrome for Testing" not in w.title})
        return _clip("\n".join(lines))
    if topic == "history":
        return _clip(recent_history(30) or "No earlier Co-opilot tasks yet.")
    if topic == "now":
        front = "\n".join(["Windows in front (top first):", *_front_windows()])
        return _clip(front + "\n" + _powershell("now", _NOW))
    if topic == "displays":
        import desktop
        return "\n".join(f"monitor {i}{' (main)' if p else ''}: {r - l}x{b - t} at ({l},{t})"
                         for i, (l, t, r, b, p) in enumerate(desktop.monitors(), 1))
    command = _COMMANDS.get(topic)
    if command is None:
        raise ValueError(f"unknown topic: {topic}")
    return _clip(_powershell(topic, command))


def _powershell(topic, command):
    # ponytail: the pipe is IBM437, which turns a curly quote into ". That breaks JSON. UTF-8 for this process.
    command = ("[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false; "
               "$OutputEncoding = [Console]::OutputEncoding\n" + command)
    try:
        done = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=25,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        return f"{topic}: timed out"
    return (done.stdout or "").strip() or (done.stderr or "").strip() or "(no output)"


def _clip(text):
    return text if len(text) <= _LIMIT else text[:_LIMIT] + "\n...(truncated)"


# Read-only diagnostics Copilot may request. The first word of each pipeline stage has to be listed.
# ponytail: a verb list, not a parser. A command that changes something is refused even if the user asks.
# Upgrade path is a real PowerShell parser if the list gets in the way of a safe check.
_PS_FIRST = {
    "get-process", "get-service", "get-computerinfo", "get-ciminstance", "get-winevent", "get-eventlog",
    "get-netadapter", "get-netipconfiguration", "get-netipaddress", "get-netroute", "get-nettcpconnection",
    "get-dnsclientserveraddress", "get-dnsclientcache", "get-pnpdevice", "get-volume", "get-disk",
    "get-physicaldisk", "get-partition", "get-hotfix", "get-itemproperty", "get-childitem", "get-item",
    "get-mpcomputerstatus", "get-mpthreatdetection", "get-bitlockervolume", "get-windowsdriver",
    "get-counter", "get-date", "resolve-dnsname", "test-netconnection", "test-connection",
}
_PS_REST = _PS_FIRST | {"measure-object", "select-object", "sort-object", "where-object", "format-table", "format-list", "out-string"}
_CMD_FIRST = {
    "ipconfig", "ping", "nslookup", "tracert", "pathping", "netstat", "route", "arp", "hostname", "whoami",
    "systeminfo", "tasklist", "driverquery", "ver", "gpresult", "dsregcmd", "nltest", "klist", "powercfg",
    "netsh", "sc", "sfc", "dism", "reg",
}
_CMD_REST = _CMD_FIRST | {"findstr", "find", "more"}
_SECRET = re.compile(r"password|credential|secret|\btoken\b|\\sam\b", re.I)
_ARGS = {
    "powercfg": re.compile(r"^powercfg(?:\s+/(?:a|availablesleepstates|q|query|lastwake|requests|waketimers|devicequery))+$", re.I),
    "sfc": re.compile(r"^sfc\s+/verifyonly$", re.I),
    "dism": re.compile(r"^dism\s+/online\s+/cleanup-image\s+/(?:checkhealth|scanhealth)$", re.I),
    "dsregcmd": re.compile(r"^dsregcmd\s+/status$", re.I),
    "klist": re.compile(r"^klist(?:\s+tickets)?$", re.I),
    "route": re.compile(r"^route\s+print$", re.I),
    "arp": re.compile(r"^arp\s+-a$", re.I),
    "sc": re.compile(r"^sc\s+(?:query|qc)(?:\s+\S+)?$", re.I),
    "reg": re.compile(r"^reg\s+query\s+\S+(?:\s+/v\s+\S+)?$", re.I),
    "gpresult": re.compile(r"^gpresult\s+/r(?:\s+/scope\s+(?:user|computer))?$", re.I),
    "nltest": re.compile(r"^nltest\s+/(?:dsgetdc:\S+|sc_query:\S+)$", re.I),
    "netsh": re.compile(r"^netsh(?:\s+\S+)*\s+show(?:\s+\S+)*$", re.I),
}


def command_ok(shell, command):
    """The command if it is one read-only check, else None."""
    command = " ".join(str(command or "").split())
    shell = str(shell or "").strip().lower()
    if shell not in ("powershell", "cmd") or not command or len(command) > 300 or _SECRET.search(command):
        return None
    if re.search(r"[;&`$<>{}\r\n]", command):
        return None
    parts = [part.strip() for part in command.split("|")]
    if not parts or any(not part for part in parts):
        return None
    first = _PS_FIRST if shell == "powershell" else _CMD_FIRST
    rest = _PS_REST if shell == "powershell" else _CMD_REST
    for index, part in enumerate(parts):
        word = part.split()[0].lower()
        if word not in (first if index == 0 else rest):
            return None
        if word == "ping" and re.search(r"(^|\s)-t(\s|$)", part, re.I):
            return None
        rule = _ARGS.get(word)
        if rule and not rule.search(part):
            return None
        if word == "netsh" and re.search(r"\b(set|add|delete|reset)\b", part, re.I):
            return None
    return command


def diagnose(command, shell, admin, stop):
    """Run one allowed check. Stop kills it. admin raises the Windows approval prompt and still cannot change anything."""
    command = command_ok(shell, command)
    if not command:
        return "Refused. That is not a read-only diagnostic, so it was not run."
    if stop is not None and stop.is_set():
        return "Stopped."
    shell = str(shell).lower()
    if admin:
        return _diagnose_admin(command, shell, stop)
    argv = (["powershell", "-NoProfile", "-NonInteractive", "-Command", command] if shell == "powershell"
            else ["cmd", "/d", "/c", command])
    proc = subprocess.Popen(argv, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            creationflags=subprocess.CREATE_NO_WINDOW)
    note = _wait(proc, stop, 20)
    if note:
        return note
    out, _err = proc.communicate()
    return _clip(f"{shell}: {command}\n{(out or '').strip() or '(no output)'}")


def _wait(proc, stop, seconds):
    """None when it finished. A short reason when Stop or the clock won."""
    deadline = time.time() + seconds
    while proc.poll() is None:
        if stop is not None and stop.is_set():
            _kill(proc)
            return "Stopped."
        if time.time() > deadline:
            _kill(proc)
            return "Timed out."
        time.sleep(0.1)
    return None


def _kill(proc):
    if proc is None or proc.poll() is not None:
        return
    _kill_pid(proc.pid)
    try:
        proc.wait(timeout=3)
    except subprocess.TimeoutExpired:
        proc.kill()


def _kill_pid(pid):
    if not pid:
        return
    subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True,
                   creationflags=subprocess.CREATE_NO_WINDOW)


def _diagnose_admin(command, shell, stop):
    import base64
    import tempfile
    path = Path(tempfile.gettempdir()) / f"coopilot-diag-{os.getpid()}.txt"
    try:
        path.unlink()
    except OSError:
        pass
    inner = command if shell == "powershell" else "cmd /d /c '" + command + "'"
    body = inner + " | Out-File -FilePath '" + str(path) + "' -Encoding utf8"
    encoded = base64.b64encode(body.encode("utf-16le")).decode()
    launcher = (
        "$p = Start-Process -FilePath powershell.exe -Verb RunAs -PassThru -WindowStyle Hidden "
        f"-ArgumentList '-NoProfile','-NonInteractive','-EncodedCommand','{encoded}'\n"
        "if (-not $p) { 'declined'; exit 1 }\n"
        "Write-Output $p.Id"
    )
    proc = subprocess.Popen(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", launcher],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
    # The Windows approval box can sit there; Stop closes our wait. The check itself then gets 20 seconds.
    note = _wait(proc, stop, 120)
    if note:
        return note
    text = (proc.stdout.read() if proc.stdout else "").strip()
    if proc.returncode != 0 or not text.isdigit():
        return "The administrator prompt was declined, so nothing ran."
    pid = int(text)
    deadline = time.time() + 20
    while True:
        if stop is not None and stop.is_set():
            _kill_pid(pid)
            return "Stopped."
        listed = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True,
                                creationflags=subprocess.CREATE_NO_WINDOW)
        if str(pid) not in (listed.stdout or ""):
            break
        if time.time() > deadline:
            _kill_pid(pid)
            return f"{shell} as administrator: {command}\nTimed out."
        time.sleep(0.2)
    try:
        written = path.read_text(encoding="utf-8", errors="replace").strip() if path.is_file() else ""
    except OSError:
        written = ""
    finally:
        try:
            path.unlink()
        except OSError:
            pass
    return _clip(f"{shell} as administrator: {command}\n{written or '(no output)'}")


def _check_diagnose():
    assert command_ok("powershell", "Get-Service Spooler") == "Get-Service Spooler"
    assert command_ok("powershell", "Get-Process | Select-Object Name, CPU")
    assert command_ok("cmd", "ipconfig /all") == "ipconfig /all"
    assert command_ok("cmd", "sfc /verifyonly") == "sfc /verifyonly"
    assert command_ok("powershell", "Remove-Item C:\\Windows") is None
    assert command_ok("powershell", "Get-Process; Remove-Item x") is None
    assert command_ok("cmd", "ping -t 8.8.8.8") is None
    assert command_ok("cmd", "dism /online /cleanup-image /restorehealth") is None
    assert command_ok("cmd", "reg delete HKLM\\x") is None
    assert command_ok("powershell", "Get-ItemProperty HKLM:\\SAM") is None
    stopped = type("S", (), {"is_set": lambda self: True})()
    assert diagnose("Get-Date", "powershell", False, stopped) == "Stopped."


def _check_history():
    global HISTORY
    import tempfile
    real, HISTORY = HISTORY, Path(tempfile.mkdtemp()) / "history.jsonl"
    try:
        assert recent_history(5) == ""
        for n in range(7):
            record_history(f"task {n}", f"answer\n{n}", n)
        lines = recent_history(5).splitlines()
        assert len(lines) == 5 and lines[0].endswith("task 2 -> answer 2") and lines[-1].endswith("task 6 -> answer 6"), lines
    finally:
        HISTORY = real


if __name__ == "__main__":
    _check_diagnose()
    _check_history()
    for name in ("system", "disks", "windows"):
        out = collect(name)
        assert out and "unknown" not in out, (name, out)
        print(f"== {name} ==\n{out}\n")
    print("sysinfo ok")
