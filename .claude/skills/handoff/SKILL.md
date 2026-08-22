---
name: handoff
description: Refresh SESSION_HANDOFF.md with current branch state and next step, and reconcile open work on the ContextBoard (board-only since 2026-08-22 — never create TODO.md). Use at end of session.
disable-model-invocation: true
---

Update one file at the repo root: `SESSION_HANDOFF.md`. Task state lives on **ContextBoard** (project `mcpOffice`, id 27) — this repo is **board-only** since 2026-08-22. **Never create `TODO.md` or `DOCS/DONE.md`**: the server refuses file-sync pushes for this project (`FileSyncClosed`), so a recreated file would sync nothing and desync the clone.

## Step 1 — gather state

Run in parallel:

- `git status`
- `git log --oneline -10`
- `git branch --show-current`

Read existing `SESSION_HANDOFF.md` so you don't drop context the user added by hand. Call `list_cards` (projectId 27) for the open work; `get_card` for any body you need.

## Step 2 — rewrite SESSION_HANDOFF.md

Use this section structure:

```
# Session Handoff — <YYYY-MM-DD>

## Where things stand

**Branch:** `<branch>` (and pushed status)
**Latest commit:** `<sha>` <subject>

## Decisions made autonomously

<only if non-trivial — design choices, deviations from the plan, things future-you needs to know that aren't obvious from the diff>

## Known nuisances

<open warnings, license wiring still pending, NU1900s, etc. — only items that are still relevant>

## What's next

<the next card(s) by AREA-NNN id with a 1–2 sentence summary, plus any prerequisite>

## How to resume

​```bash
cd C:/Projects/mcpOffice
git status
dotnet build
dotnet test
​```
```

## Step 3 — reconcile the board

- Work finished this session: `complete_card` with a `conclusion` that cites the closing commit SHA(s).
- New follow-ups raised this session: `add_card` (project 27) with an `AREA-NNN:` title, a type tag, and a body that carries the context (file refs, evidence, plan). Refer to cards by that id in the handoff.
- Bodies that changed: `update_card`. Don't retitle or move cards here unless asked.

## Bounds

Don't change anything else in the repo. No code edits, no plan-doc edits, no commits. Rewrite the handoff, reconcile the board, and report what changed in 2–3 lines.
