---
name: handoff
description: Refresh SESSION_HANDOFF.md and TODO.md with current branch state, completed plan tasks, and next step. Use at end of session.
disable-model-invocation: true
---

Update two files at the repo root: `SESSION_HANDOFF.md` and `TODO.md`. Create `TODO.md` if it doesn't exist.

## Step 1 — gather state

Run in parallel:

- `git status`
- `git log --oneline -10`
- `git branch --show-current`

Read `docs/plans/2026-04-30-mcpoffice-word-poc-plan.md` to identify which tasks are ✅ (already done in commits/code) and which is the next ⬜.

Read existing `SESSION_HANDOFF.md` and `TODO.md` (if present) so you don't drop context the user added by hand.

## Step 2 — rewrite SESSION_HANDOFF.md

Use this section structure:

```
# Session Handoff — <YYYY-MM-DD>

## Where things stand

**Branch:** `<branch>` (and pushed status)
**Latest commit:** `<sha>` <subject>

Plan tasks:
✅ Task N — <one-liner>
⬜ Task M — <one-liner>
…

## Decisions made autonomously

<only if non-trivial — design choices, deviations from the plan, things future-you needs to know that aren't obvious from the diff>

## Known nuisances

<open warnings, license wiring still pending, NU1900s, etc. — only items that are still relevant>

## What's next

<the next ⬜ task with a 1–2 sentence summary, plus any prerequisite>

## How to resume

​```bash
cd C:/Projects/mcpOffice
git status
dotnet build
dotnet test
​```
```

## Step 3 — maintain TODO.md

If TODO.md doesn't exist, create it:

```
# TODO

Pending work for mcpOffice. Pulled from docs/plans/2026-04-30-mcpoffice-word-poc-plan.md plus session-level items.

## Plan tasks

- [x] Task 1 — …
- [ ] Task N — …

## Side items

- [ ] <e.g., wire DevExpress runtime license via licenses.licx>
- [ ] <e.g., drop DevExpress feed from nuget.config>
```

If TODO.md already exists:

- Update checkbox state to match reality (`[ ]` → `[x]` for tasks now done).
- Add new pending items raised in this session under **Side items** if they aren't already in the plan.
- Don't delete completed items in the same session unless asked — they show momentum.

## Bounds

Don't change anything else in the repo. No code edits, no plan-doc edits, no commits. Just rewrite these two files and report what changed in 2–3 lines.
