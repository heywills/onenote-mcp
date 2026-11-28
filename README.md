# OneNote MCP server

MCP server that reads, searches, creates, and deletes pages in every notebook open in OneNote desktop, including SharePoint and OneDrive notebooks. It drives the desktop app through COM, so it uses the app's own sign-in: no Entra app, no API keys.

Forked from [wgthomas/onenote-mcp](https://github.com/wgthomas/onenote-mcp) (MIT) and ported from Python to .NET.

## Tools

| Tool | Purpose |
|---|---|
| `onenote_list_notebooks` | Open notebooks |
| `onenote_list_sections` | Sections in a notebook, including section groups |
| `onenote_list_pages` | Pages in a section |
| `onenote_get_notebook_tree` | Notebooks, groups, sections, and page titles |
| `onenote_get_page` | Page as Markdown or HTML: headings, nested lists, tables, code blocks, inline formatting, and images |
| `onenote_get_page_by_link` | Page from a "Copy link to page" link (`onenote:` or web link) |
| `onenote_get_page_raw` | Page as OneNote XML |
| `onenote_get_page_images` | Images on a page |
| `onenote_get_image` | One image by callback ID |
| `onenote_search` | Search all notebooks |
| `onenote_search_in_notebook` | Search one notebook |
| `onenote_create_page` | New page from HTML: headings, nested lists, tables, code, images. Reads back the HTML `onenote_get_page` saves. |
| `onenote_create_page_from_markdown` | New page from Markdown. Reads back as the same Markdown, images included. |
| `onenote_delete_page` | Moves a page to Deleted Pages. Subpages stay. |

`onenote_get_page`, `onenote_get_page_by_link`, and `onenote_get_page_raw` save to `outputPath`, an absolute file path the agent names, and return the path. The Markdown tools save images full size to a folder next to the page and link them: `page.md` gets `page_images`. `inline=true` returns the page in the response instead, with `[Image N]` and callback IDs for the image tools.

`format=html` keeps what Markdown drops:
- Colors, highlights, fonts, and sizes
- Alignment and indented paragraphs
- Table header rows and cell shading
- Image sizes

`onenote_create_page` reads it back, with the title from `<title>`. Tags, ink, and paragraph spacing aren't kept. OneNote also rewrites some markup (`#FF0000` becomes `red`) and drops trailing spaces.

`onenote_create_page` and `onenote_create_page_from_markdown` take the body inline or from `inputPath`, an absolute UTF-8 file path. They embed local images:
- From a file: only images inside the file's folder, by relative path
- Inline: absolute paths too
- Web URLs, network shares and mapped network drives, links that lead out of the file's folder or to a network, files over 50 MB, and other types stay links. So do all images of a file on a share.
- An image inside a heading, a link, formatting, or a table header stays a link.
- Images over 10 MB are shrunk. Images that would push a page past 50 MB stay links.

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

- Search matches titles holding every query word, then text through Windows Search (indexed pages only). Skips deleted pages. A query with `OR`, `NOT`, `-word`, or parentheses searches text only.
- Images OneNote hasn't downloaded yet return an error from the image tools, and show as `[Image N: reason]` in saved pages. Open the page in OneNote desktop, then retry.
- Markdown that OneNote can't hold:
  - Indented paragraphs outside a list come back flush, because indented Markdown reads as code.
  - Numbered lists start at 1. OneNote numbers lists itself, so separate lists alternate between `1.` and `1)` in OneNote to restart.
  - A hard line break splits a paragraph in two, and becomes a space in a list item. OneNote drops line breaks inside a paragraph.
  - An image in a paragraph splits it: text, image, text. In a list item, it moves below the item text. Code blocks lose their language, and OneNote stores a tab in code as 8 spaces.
  - Headings, tables, and code inside a quote lose the quote. Nested quotes flatten to one level.
  - Bold that covers a whole table header cell is dropped, since the header row is already bold. So is italic that covers a whole h4 or h6 heading, since those styles are italic.

## Tests

```powershell
dotnet test
```

Live tests run against OneNote desktop when `ONENOTE_LIVE_SECTION` is set. They create pages in that section and delete them permanently. The delete test leaves its page in Deleted Pages, because OneNote refuses deletes from there.

```powershell
$env:ONENOTE_LIVE_SECTION = "My Notebook/Test section"; dotnet test
```

Set `ONENOTE_LIVE_IMAGE_LINK` to a "Copy link to page" link to also test images.

## How it works

- Calls OneNote through `Application2Class` in OneNote's interop assembly, which Office installs. Late binding fails in 64-bit .NET because OneNote registers its type library for 32-bit only.
- The build references that assembly from the GAC, so it builds only where OneNote desktop is installed.
- COM errors come back as tool errors with OneNote's error name and code.

## License

MIT
