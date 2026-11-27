# OneNote MCP server

MCP server that reads, searches, and creates pages in every notebook open in OneNote desktop, including SharePoint and OneDrive notebooks. It drives the desktop app through COM, so it uses the app's own sign-in: no Entra app, no API keys.

Forked from [wgthomas/onenote-mcp](https://github.com/wgthomas/onenote-mcp) (MIT) and ported from Python to .NET.

## Tools

| Tool | Purpose |
|---|---|
| `onenote_list_notebooks` | Open notebooks |
| `onenote_list_sections` | Sections in a notebook, including section groups |
| `onenote_list_pages` | Pages in a section |
| `onenote_get_notebook_tree` | Notebooks, groups, sections, and page titles |
| `onenote_get_page` | Page as Markdown: headings, nested lists, tables, code blocks, inline formatting, and image callback IDs |
| `onenote_get_page_by_link` | Page from a "Copy link to page" link (`onenote:` or web link) |
| `onenote_get_page_raw` | Page as OneNote XML |
| `onenote_get_page_images` | Images on a page |
| `onenote_get_image` | One image by callback ID |
| `onenote_search` | Search all notebooks |
| `onenote_search_in_notebook` | Search one notebook |
| `onenote_create_page` | New page from HTML: headings, nested lists, tables, code |
| `onenote_create_page_from_markdown` | New page from Markdown. Reads back as the same Markdown. |

## Requirements

- Windows with OneNote desktop (Microsoft 365), not the Store version. COM starts OneNote if it's closed.
- .NET 10 SDK

## Setup

1. Build:
   ```powershell
   dotnet publish src/OneNoteMcp -c Release -o publish
   ```
2. Register at user scope:
   ```powershell
   claude mcp add onenote --scope user -- "C:/path/to/onenote-mcp/publish/OneNoteMcp.exe"
   ```
3. Restart Claude Code and ask: `Use the onenote MCP server to list my notebooks.`

## Notes

- Search uses Windows Search, so it finds indexed pages only.
- Images OneNote hasn't downloaded yet return an error. Open the page in OneNote desktop, then retry.
- Indented paragraphs outside a list come back flush, because indented Markdown reads as code.

## How it works

- Calls OneNote through `Application2Class` in OneNote's interop assembly, which Office installs. Late binding fails in 64-bit .NET because OneNote registers its type library for 32-bit only.
- The build references that assembly from the GAC, so it builds only where OneNote desktop is installed.
- COM errors come back as tool errors with OneNote's error name and code.

## License

MIT
