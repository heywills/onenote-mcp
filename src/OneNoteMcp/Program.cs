using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OneNoteMcp;

var builder = Host.CreateApplicationBuilder(args);

// stdout carries the MCP protocol, so logs go to stderr.
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services
    .AddSingleton<OneNoteClient>()
    .AddMcpServer(o => o.ServerInstructions =
        "Access OneNote desktop notebooks through COM. Read, search, and create pages, and see embedded images.")
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
