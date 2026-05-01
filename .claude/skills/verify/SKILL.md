---
name: verify
description: Run dotnet build + dotnet test and report a clean pass/fail. Use before claiming a task done.
---

Verification gate. Run these two commands in order at the repo root:

1. `dotnet build --nologo`
2. `dotnet test --nologo --logger "console;verbosity=normal"`

Reporting rules:

- ✅ if both succeed. The 6× NU1900 warnings on build are expected (see CLAUDE.md) — ignore them, they don't count as failures.
- ❌ if either fails. Report:
  - which step failed (build vs test)
  - the failing test names (for test failures) or the first 1–2 compiler errors (for build failures)
  - nothing else — no diagnosis, no proposed fix, no narration

This skill is a gate, not a debug session. Don't propose fixes unless the user asks. Don't run `dotnet restore` or other commands unless the build/test output explicitly demands it.
