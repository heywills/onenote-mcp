---
description: Don't add code metric warnings (CA1501, CA1502, CA1505, CA1506, CA1508)
paths:
  - "src/**/*.cs"
---

# Code Metrics

`.editorconfig` reports these as warnings. Changes must not add to the count.

- CA1502 cyclomatic complexity > 25 per method
- CA1505 maintainability index < 10
- CA1506 class coupling > 95 per type, > 40 per member
- CA1501 inheritance depth > 5
- CA1508 dead conditional code

## Check

`dotnet build OneNoteMcp.slnx --no-incremental`

- Incremental builds skip warnings from projects that don't recompile

## Fixing

- Extract methods or classes; don't suppress (`#pragma`, `[SuppressMessage]`, `.editorconfig`) without asking
- Each `?.` and `??` adds a branch; after a null guard, use `.` (also fixes CA1508)
- `partial` classes don't reduce CA1506; the analyzer counts all parts as one type
