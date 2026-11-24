# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Windows desktop app in Python. The window is HTML/CSS in Microsoft Edge WebView2 through pywebview (chosen by the user for presentation). The click-through "hands off" overlay stays native tkinter. No web server, no external fonts or CDNs: everything ships with the app.

## Users

Ontario Public Service staff across the ministry, from developers to policy and program staff, with very different comfort levels with computers and AI. They use it at their desk between meetings: to read their own calendar, mail, and Teams meeting transcripts, to diagnose a slow or broken PC, and to have the PC do small tasks (open an app, type, click) for them.

## Product Purpose

Co-opilot is a side panel that answers through the user's own signed-in Microsoft 365 Copilot (no separate AI service) and can act on the PC: invisible reads first, mouse and keyboard only when a task needs the screen. Success is a task done correctly in one ask, with the user always knowing what is happening and able to stop it.

## Positioning

It uses the work Copilot the user already has, inside government sign-in, and it can actually operate this Windows PC and read their Outlook and Teams without screen scraping for reads. A neighbouring chatbot can answer; it cannot do the task on this machine.

## Operating Context

- Runs beside other work on one or two monitors; the window is narrow (about 520 px) and sits on the right.
- While it drives the mouse, the window hides and a translucent overlay with a Stop button covers every monitor. Moving the mouse to a screen corner also stops it.
- It asks before risky actions (terminals, deletes, sending) with a question card in the conversation.
- Account answers (meetings, mail, tasks, files, recordings) come from Outlook and Teams in the background.

## Capabilities and Constraints

- Conversation with markdown-light answers: headings, bullets, bold, inline code.
- Question cards: single or multiple choice plus optional free text, Submit and Skip. Skip ends the task.
- Status: starting, preparing chat, sign-in, ready, working (step n), using the mouse and keyboard, done, stopped, error.
- New chat clears memory. A developer log opens with Ctrl+Shift+D.
- No telemetry, no external network calls from the UI.

## Brand Commitments

Name is "Co-opilot" (that spelling). Government of Ontario setting: plain, trustworthy, calm; no hype.

## Evidence on Hand

No user research, testimonials, or metrics exist. Do not invent any.

## Product Principles

1. Say what is happening, always: status is visible, and every action has visible feedback.
2. The user stays in control: Stop is one press away; risky actions ask first.
3. Answer first: the first sentence is the fact, details follow.
4. Quiet by default: nothing competes with the conversation.

## Accessibility & Inclusion

Ministry-wide rollout under AODA: WCAG 2.2 AA is the floor, AAA contrast for body text where feasible. Full keyboard use, visible focus, screen reader announcements for new replies and status changes, reduced-motion respected, Windows light and dark themes, text that scales with Windows display scaling.
