"""File skills: list folders, read documents (text, CSV, Excel, Word), and write text, CSV, or Excel files."""

import csv
import html
import io
import os
import re
import time
import zipfile
from pathlib import Path

import openpyxl

from desktop import _RUNNABLE

READ_LIMIT = 6000
_TEXT = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".log", ".ini", ".cfg", ".yaml", ".yml",
         ".py", ".js", ".ts", ".ps1", ".sql", ".cs", ".java", ".css", ".bat", ".cmd", ".sh", ".toml", ".env"}


def resolve(path):
    """Expand %VARS% and ~; relative paths start in the user's home folder."""
    target = Path(os.path.expandvars(os.path.expanduser(str(path).strip().strip('"'))))
    return target if target.is_absolute() else Path.home() / target


def list_dir(path):
    folder = resolve(path)
    if not folder.is_dir():
        raise ValueError(f"not a folder: {folder}")
    entries = sorted(folder.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
    lines = [f"{folder} ({len(entries)} items)"]
    for entry in entries[:200]:
        try:
            stat = entry.stat()
        except OSError:
            continue
        when = time.strftime("%Y-%m-%d %H:%M", time.localtime(stat.st_mtime))
        lines.append(f"  {entry.name}/  {when}" if entry.is_dir() else f"  {entry.name}  {_size(stat.st_size)}  {when}")
    if len(entries) > 200:
        lines.append(f"  ...and {len(entries) - 200} more")
    return "\n".join(lines)


def read_file(path, sheet=None):
    target = resolve(path)
    if not target.is_file():
        raise ValueError(f"file not found: {target}")
    suffix = target.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        text = _read_excel(target, sheet)
    elif suffix == ".docx":
        text = _read_docx(target)
    elif suffix in _TEXT or target.stat().st_size < 200_000 and _looks_text(target):
        text = target.read_text(encoding="utf-8-sig", errors="replace")
    else:
        raise ValueError(f"cannot read {suffix or 'this'} files as text; open it with open_path instead")
    header = f"{target} ({_size(target.stat().st_size)})\n"
    return header + (text if len(text) <= READ_LIMIT else text[:READ_LIMIT] + f"\n...(truncated, {len(text)} chars total)")


def excel_rows(path, sheet=None):
    """Rows of a sheet as dicts keyed by the header row; blank rows skipped."""
    book = openpyxl.load_workbook(resolve(path), read_only=True, data_only=True)
    try:
        ws = book[sheet] if sheet else book.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else "" for h in next(rows, [])]
        out = []
        for row in rows:
            item = {header[i]: _cell(v) for i, v in enumerate(row) if i < len(header) and header[i]}
            if any(item.values()):
                out.append(item)
        return out
    finally:
        book.close()


def write_file(path, content=None, rows=None, mode="create"):
    """Create, append to, or overwrite a file. .xlsx and .csv take rows (lists); others take text."""
    target = resolve(path)
    if target.suffix.lower() in _RUNNABLE:
        raise ValueError(f"refusing to write {target.suffix} files")
    if mode == "create" and target.exists():
        raise ValueError(f"already exists: {target}. Use mode append or overwrite.")
    target.parent.mkdir(parents=True, exist_ok=True)
    suffix = target.suffix.lower()
    if suffix == ".xlsx":
        if mode == "append" and target.exists():
            book = openpyxl.load_workbook(target)
            ws = book.active
        else:
            book = openpyxl.Workbook()
            ws = book.active
        for row in rows or _text_rows(content):
            ws.append([_excel_value(v) for v in row])
        book.save(target)
    elif suffix == ".csv" and rows is not None:
        with target.open("a" if mode == "append" else "w", newline="", encoding="utf-8-sig") as file:
            csv.writer(file).writerows(rows)
    else:
        text = content if content is not None else "\n".join("\t".join(map(str, r)) for r in rows or [])
        with target.open("a" if mode == "append" else "w", encoding="utf-8") as file:
            file.write(text)
    return f"{'appended to' if mode == 'append' else 'wrote'} {target} ({_size(target.stat().st_size)})"


def _read_excel(target, sheet):
    book = openpyxl.load_workbook(target, read_only=True, data_only=True)
    try:
        sheets = [book[sheet]] if sheet else book.worksheets
        parts = [f"Sheets: {', '.join(book.sheetnames)}"]
        for ws in sheets:
            parts.append(f"== {ws.title} ==")
            for index, row in enumerate(ws.iter_rows(values_only=True)):
                if index >= 300:
                    parts.append("...(more rows)")
                    break
                if any(v is not None for v in row):
                    parts.append("\t".join(_cell(v) for v in row).rstrip())
        return "\n".join(parts)
    finally:
        book.close()


def _read_docx(target):
    with zipfile.ZipFile(target) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"<w:tab/>", "\t", xml)
    return html.unescape(re.sub(r"<[^>]+>", "", xml)).strip()


def _looks_text(target):
    return b"\x00" not in target.read_bytes()[:4096]


def _text_rows(content):
    return list(csv.reader(io.StringIO(content or ""), delimiter="\t" if "\t" in (content or "") else ","))


def _cell(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _excel_value(value):
    # Leading zeros ("007", IDs, phone numbers) stay text.
    if isinstance(value, str) and re.fullmatch(r"-?(0|[1-9]\d*)(\.\d+)?", value.strip()):
        return float(value) if "." in value else int(value)
    return value


def _size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def _check():
    import tempfile
    folder = Path(tempfile.mkdtemp())
    sheet = folder / "items.xlsx"
    write_file(sheet, rows=[["Title", "Acceptance Criteria", "Tags"], ["Login page", "Given a user\nWhen...", "ui; auth"],
                            ["", "", ""], ["Logout", "Session ends", "auth"]])
    rows = excel_rows(sheet)
    assert rows == [{"Title": "Login page", "Acceptance Criteria": "Given a user\nWhen...", "Tags": "ui; auth"},
                    {"Title": "Logout", "Acceptance Criteria": "Session ends", "Tags": "auth"}], rows
    assert "Login page" in read_file(sheet) and "== Sheet ==" in read_file(sheet)
    note = folder / "note.txt"
    write_file(note, "hello")
    try:
        write_file(note, "again")
        raise AssertionError("create must not overwrite")
    except ValueError:
        pass
    write_file(note, " world", mode="append")
    assert read_file(note).endswith("hello world")
    try:
        write_file(folder / "x.bat", "echo")
        raise AssertionError("runnable files must be refused")
    except ValueError:
        pass
    assert "items.xlsx" in list_dir(folder)
    print("files ok")


if __name__ == "__main__":
    _check()
