using Microsoft.Office.Interop.OneNote;
using ModelContextProtocol;

namespace OneNoteMcp.Tests;

/// <summary>
/// Stands in for OneNote. Each call runs a handler a test sets, and fails like OneNote until then.
/// </summary>
internal sealed class FakeOneNote : IOneNoteClient
{
    public const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public Func<string, HierarchyScope, string> Hierarchy { get; set; } = (_, _) => throw NotSet();

    public Func<string, string> PageContent { get; set; } = _ => throw NotSet();

    public Func<string, string, byte[]> Image { get; set; } = (_, _) => throw NotSet();

    public Func<string, string, string> Find { get; set; } = (_, _) => throw NotSet();

    public Func<string, string> NewPage { get; set; } = _ => throw NotSet();

    public Action<string> Update { get; set; } = _ => { };

    public Action<string, bool> Delete { get; set; } = (_, _) => { };

    /// <summary>Every DeleteHierarchy call, in order</summary>
    public List<(string Id, bool Permanently)> Deleted { get; } = [];

    public string GetHierarchy(string startNodeId, HierarchyScope scope) => Hierarchy(startNodeId, scope);

    public string GetPageContent(string pageId) => PageContent(pageId);

    public byte[] GetBinaryPageContent(string pageId, string callbackId) => Image(pageId, callbackId);

    public string FindPages(string query, string startNodeId = "") => Find(query, startNodeId);

    public string GetHyperlink(string objectId) => throw NotSet();

    public string CreateNewPage(string sectionId) => NewPage(sectionId);

    public void UpdatePageContent(string xml) => Update(xml);

    public void DeleteHierarchy(string objectId, bool permanently = false)
    {
        Deleted.Add((objectId, permanently));
        Delete(objectId, permanently);
    }

    /// <summary>A page with a title and these OEs.</summary>
    public static string Page(string oes = "") => $"""
        <one:Page xmlns:one="{Ns}" ID="page" name="P">
          <one:Title><one:OE><one:T><![CDATA[P]]></one:T></one:OE></one:Title>
          <one:Outline><one:OEChildren>{oes}</one:OEChildren></one:Outline>
        </one:Page>
        """;

    private static McpException NotSet() => new("Not set in this test");
}
