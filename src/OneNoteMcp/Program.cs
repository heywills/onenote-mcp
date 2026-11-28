using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OneNoteMcp;

var builder = Host.CreateApplicationBuilder(args);

// stdout carries the MCP protocol, so logs go to stderr.
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services
    .AddSingleton<IOneNoteClient, OneNoteClient>()
    .AddMcpServer(o => o.ServerInstructions =
        "Access OneNote desktop notebooks through COM. Read, search, create, and delete pages, and see embedded images. onenote_get_page, onenote_get_page_by_link, and onenote_get_page_raw save to a file you name unless you pass inline=true. onenote_create_page and onenote_create_page_from_markdown read the body from inputPath, an absolute file path, or inline.")
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
