using Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp;

/// <summary>
/// The OneNote calls the tools make, so tests can stand in for OneNote.
/// </summary>
/// <remarks>
/// - OneNote failures are <see cref="ModelContextProtocol.McpException"/>.
/// </remarks>
public interface IOneNoteClient
{
    /// <summary>Hierarchy XML from <paramref name="startNodeId"/> down to <paramref name="scope"/>. Empty for all notebooks.</summary>
    string GetHierarchy(string startNodeId, HierarchyScope scope);

    /// <summary>A page's XML.</summary>
    string GetPageContent(string pageId);

    /// <summary>An image's bytes. Throws when OneNote returns no data or unreadable data.</summary>
    byte[] GetBinaryPageContent(string pageId, string callbackId);

    /// <summary>Windows Search's matching pages, as hierarchy XML.</summary>
    string FindPages(string query, string startNodeId = "");

    /// <summary>The <c>onenote:</c> link to a page, section, or notebook.</summary>
    string GetHyperlink(string objectId);

    /// <summary>A new blank page with a title. Returns its ID.</summary>
    string CreateNewPage(string sectionId);

    /// <summary>Replaces a page's content with this page XML.</summary>
    void UpdatePageContent(string xml);

    /// <summary>Deletes a page, section, or notebook.</summary>
    /// <param name="permanently">Skip the notebook's recycle bin</param>
    void DeleteHierarchy(string objectId, bool permanently = false);
}
