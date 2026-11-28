using Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp;

/// <summary>
/// The OneNote calls the tools make, so tests can stand in for OneNote.
/// </summary>
/// <remarks>
/// - Errors are <see cref="ModelContextProtocol.McpException"/>, with OneNote's error name.
/// </remarks>
public interface IOneNoteClient
{
    string GetHierarchy(string startNodeId, HierarchyScope scope);

    string GetPageContent(string pageId);

    /// <summary>An image's bytes. Throws when OneNote returns no data or unreadable data.</summary>
    byte[] GetBinaryPageContent(string pageId, string callbackId);

    string FindPages(string query, string startNodeId = "");

    string GetHyperlink(string objectId);

    string CreateNewPage(string sectionId);

    void UpdatePageContent(string xml);

    /// <param name="permanently">Skip the notebook's recycle bin</param>
    void DeleteHierarchy(string objectId, bool permanently = false);
}
