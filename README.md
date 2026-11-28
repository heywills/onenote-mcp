# OneNote MCP server

An MCP server that lets AI agents read, search, create, and delete pages in OneNote desktop notebooks.

## Background

Most OneNote MCP servers use the Microsoft Graph API. That works, but it's hard for agents in practice:
- Graph needs an Entra app registration and user consent, and some organizations restrict both.
- Personal notebooks, SharePoint site notebooks, and Teams group notebooks sit under different API paths, so a server has to find and query each one.
- Responses carry a lot of metadata unless trimmed with [`$select`](https://learn.microsoft.com/en-us/graph/onenote-best-practices), which fills an agent's context.
- Page images are API resources that need an authenticated download, so an agent can't open them from the page HTML.

Returning a whole page in a tool response also loses its images and fills the agent's context.

## Solution

This server drives OneNote desktop through its local COM API instead of Graph, so it uses the app's own sign-in:

| Graph API | This server |
|---|---|
| App registration and consent | OneNote desktop's own sign-in |
| Different paths for personal, SharePoint, and Teams notebooks | Every notebook open in OneNote desktop |
| Metadata-heavy responses | Short hierarchy summaries, and pages saved to files |
| Images behind authenticated downloads | Image files saved next to the page, or MCP images |

It's built for agents with file access:
- Page reads save to a Markdown or HTML file, with the page's images in a folder next to it.
- Page creates read a Markdown or HTML file, and embed its local images.
- A saved page comes back nearly the same when created again (see [Limits](#limits)), so an agent can save a page, edit the file, and create the result.

Forked from [wgthomas/onenote-mcp](https://github.com/wgthomas/onenote-mcp) (MIT) and ported from Python to .NET.

## Compatibility

- Windows with OneNote desktop from Microsoft 365. The Store version of OneNote has no COM API.
- .NET 10 SDK, to build
- Any MCP client that runs stdio servers. The steps below use Claude Code.

## Installation

1. Clone the repo and build:
   ```powershell
   git clone https://github.com/heywills/onenote-mcp.git
   cd onenote-mcp
   dotnet publish src/OneNoteMcp -c Release -o publish
   ```
   The build references OneNote's interop assembly, which Office installs, so it builds only where OneNote desktop is installed.
2. Register the server at user scope:
   ```powershell
   claude mcp add onenote --scope user -- "C:/path/to/onenote-mcp/publish/OneNoteMcp.exe"
   ```
   For another client, add the exe as a stdio server. For example, in Claude Desktop or Cursor:
   ```json
   { "mcpServers": { "onenote": { "command": "C:/path/to/onenote-mcp/publish/OneNoteMcp.exe" } } }
   ```
   VS Code uses a `servers` key, with `"type": "stdio"` on the server.
3. Restart the client and ask: `Use the onenote MCP server to list my notebooks.`

OneNote starts on the first call if it's closed.

## Usage

Ask in plain language. For example:
- `Save my "Weekly 1:1" page to C:\work\one-on-one.md`
- `Find my notes about the website launch`
- `Create a page in Work/Projects from C:\work\plan.md`
- `Save this page as HTML, fix the typos, and create it as a new page`

### Tools

| Tool | Purpose |
|---|---|
| `onenote_list_notebooks` | Open notebooks |
| `onenote_list_sections` | Sections in a notebook, including section groups |
| `onenote_list_pages` | Pages in a section |
| `onenote_get_notebook_tree` | Notebooks, groups, sections, and page titles |
| `onenote_search` | Search all notebooks |
| `onenote_search_in` | Search one notebook, section group, or section |
| `onenote_get_page` | Save a page as Markdown or HTML |
| `onenote_get_page_by_link` | Save a page from a "Copy link to page" link: the `onenote:` link, the web link, or both |
| `onenote_get_page_raw` | Save a page as OneNote XML, for debugging |
| `onenote_get_page_images` | Images on a page, for the agent to see |
| `onenote_get_image` | One image, by callback ID |
| `onenote_create_page` | New page from HTML |
| `onenote_create_page_from_markdown` | New page from Markdown |
| `onenote_move_page` | Move a page and its subpages to another section or notebook, or to another place in its section |
| `onenote_delete_page` | Move a page to Deleted Pages |

### Reading pages

The read tools save to `outputPath`, an absolute file path the agent names, and return the path. Images go full size into a folder next to the page: `page.md` gets `page_images`. Pass `inline=true` to get the page in the response instead, with callback IDs for the image tools.

Markdown keeps headings, nested lists, tables, code, inline formatting, links, and images. `format=html` also keeps:
- Colors, highlights, fonts, and sizes
- Alignment and indented paragraphs
- Table header rows and cell shading
- Image sizes
- Where each note container sits on the page, so a margin note stays beside its text. Each is a `div class="outline"`: to edit such a page, add text inside a div, or it lands over the main text. An image placed freely on such a page comes back inside a note container of its own.

An image OneNote hasn't downloaded yet shows as `[Image N: reason]`, with any text OneNote recognized in it. Open the page in OneNote desktop, then save it again.

### Creating pages

The create tools take the body inline or from `inputPath`, an absolute UTF-8 file path. The title comes from the `title` parameter, the Markdown's leading `# ` heading, or the HTML's `<title>`.

Local PNG, JPEG, GIF, BMP, and TIFF images are embedded:
- From a file: images inside the file's folder, by relative path
- Inline: absolute paths only
- Images over 10 MB are shrunk.

These stay links, and the result counts both:
- Web images
- Network shares and mapped network drives, and every image of a file on a share
- Files outside the folder, and links that lead out of it
- Files over 50 MB, and images that would push the page past 50 MB
- Images inside headings, links, formatting, or table headers

### Tags

Pages keep their OneNote tags:
- Checkbox tags (To Do, Attendance) are task items in Markdown: `- [ ]` and `- [x]`. Creating from Markdown turns task items into To Do checkboxes.
- Other tags (Important, Question) are labels in brackets, such as `[Important]`. Creating from Markdown keeps them as text.
- HTML keeps tags with `data-tag` attributes, so `onenote_create_page` restores them.

### Search

Search matches page titles that contain every query word, then page text through Windows Search, which finds indexed pages only. Title matches come first, then text matches, each newest first, up to `maxResults` (20).
- Title search finds new pages that Windows Search hasn't indexed yet.
- A query with `OR`, `NOT`, `-word`, `+word`, or a property filter such as `author:x` searches text only.
- Deleted pages are skipped.

### Moving pages

`onenote_move_page` moves a page and its subpages to a section in any notebook, at the end or after `afterPageId` as that page's sibling. Pass the page's own section to reorder it. Pages that change section get new IDs, and the result lists each old and new ID. Their "Copy link to page" page-id stays the same.

### Deleting pages

`onenote_delete_page` moves a page to its notebook's Deleted Pages, where OneNote keeps it for 60 days. It refuses sections and notebooks. The page's subpages stay in the section.

## Limits

Markdown can't hold everything a OneNote page can:
- Indented paragraphs outside a list come back flush, because indented Markdown reads as code.
- An image in a paragraph splits it: text, image, text. In a list item, it moves below the item text.
- A hard line break splits a paragraph in two, and becomes a space in a list item. OneNote drops line breaks inside a paragraph.
- Headings, tables, and code inside a quote lose the quote. Nested quotes flatten to one level.
- Code blocks lose their language, and OneNote stores a tab in code as 8 spaces.
- Note containers come one after another, so a margin note follows the main text.
- Numbered lists start at 1. Separate lists alternate between `1.` and `1)` in OneNote, so each restarts.
- Bold covering a whole table header cell is dropped, since the header row is already bold. So is italic covering a whole h4 or h6 heading.
- Tags in table cells, checkboxes on headings, labels on empty paragraphs, and a To Do's bullet don't survive.

Both formats drop ink and paragraph spacing, and tags on plain code lines and on paragraphs that hold only an image. OneNote also rewrites some markup (`#FF0000` becomes `red`) and drops trailing spaces.

## Troubleshooting

- **"Section not found. Is the notebook open in OneNote desktop?"**: open the notebook in OneNote, then retry.
- **An image tool says the image isn't downloaded**: open the page in OneNote desktop so it downloads, then retry.
- **An error ends with a code, such as "The page is read-only (0x8004200C)"**: it comes from OneNote's COM API. That one means the page is in Deleted Pages.
- **`dotnet publish` can't copy files**: the MCP client is running the server, and Windows locks the exe. Close the client, or publish to another folder and point the client at it.

## How it works

- Calls OneNote through `Application2Class` in OneNote's interop assembly. Late binding fails in 64-bit .NET, because OneNote registers its type library for 32-bit only.
- COM calls run one at a time, as OneNote requires.
- Converts between OneNote's page XML and Markdown (Markdig) or HTML (AngleSharp).

## Contributing

Run the tests:

```powershell
dotnet test
```

Live tests run against OneNote desktop when `ONENOTE_LIVE_SECTION` names a test section. They create pages there and delete them permanently. The delete test leaves its page in Deleted Pages, because OneNote refuses deletes from there.

```powershell
$env:ONENOTE_LIVE_SECTION = "My Notebook/Test section"; dotnet test
```

Set `ONENOTE_LIVE_IMAGE_LINK` to a "Copy link to page" link for a page with images to also test images.

The build treats code complexity analyzer findings as warnings. Keep it at 0 warnings.

## License

MIT
