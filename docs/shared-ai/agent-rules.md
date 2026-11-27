---
description: How to add or change agent rule files for Claude and Copilot
paths:
  - "**/*.md"
---

# Agent Rules

Each rule is one shared doc plus two pointer files.

| File | Holds |
| ---- | ----- |
| `docs/shared-ai/<name>.md` | The rule, with `description` and `paths` frontmatter |
| `.claude/rules/<name>.md` | The same frontmatter and `@../../docs/shared-ai/<name>.md` |
| `.github/instructions/<name>.instructions.md` | `applyTo` and a Markdown link to the shared doc |

## Adding a rule

- Write the rule in the shared doc only.
- Scope it to the files it covers. No `paths` (Claude) or `applyTo: "**"` (Copilot) loads it in every session, so use that only for a rule every task needs.
- Keep `paths` and `applyTo` matching. `description` goes in the shared doc and the Claude pointer.

## Changing a rule

- Edit the shared doc. Touch the pointer files only to change the scope or name.
- Rename all three files together.

## Gotchas

- Claude: one `paths` entry per file. A file with more stops loading. Use brace expansion instead: `"{src,tests}/**/*.cs"`.
- Claude: `@` imports a `.md` file no matter what the parent's `paths` says. The shared doc's own frontmatter scopes it, so keep it.
- Claude: `@` skips files that don't end in `.md`.
- Copilot: keep `.github/instructions/` flat. Nested folders may not load.
