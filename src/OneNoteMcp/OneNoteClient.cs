using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;
using ModelContextProtocol;

namespace OneNoteMcp;

/// <summary>
/// Calls the OneNote desktop app through COM.
/// </summary>
/// <remarks>
/// - Uses <c>Application2Class</c> from Office's interop assembly. Late binding fails in 64-bit .NET because OneNote registers its type library for 32-bit only.
/// - Calls are serialized: OneNote handles one COM call at a time.
/// - COM errors become <see cref="McpException"/> with the OneNote error name, so tools report the real cause.
/// </remarks>
public sealed class OneNoteClient
{
    private readonly Lock _lock = new();
    private Application2Class? _app;

    public string GetHierarchy(string startNodeId, HierarchyScope scope) =>
        Call(app => { app.GetHierarchy(startNodeId, scope, out string xml, XMLSchema.xs2013); return xml; });

    public string GetPageContent(string pageId) =>
        Call(app => { app.GetPageContent(pageId, out string xml, PageInfo.piBasic, XMLSchema.xs2013); return xml; });

    public byte[] GetBinaryPageContent(string pageId, string callbackId)
    {
        var base64 = Call(app => { app.GetBinaryPageContent(pageId, callbackId, out string b64); return b64; });
        byte[] data;
        try
        {
            data = Convert.FromBase64String(base64);
        }
        catch (Exception e) when (e is FormatException or ArgumentNullException)
        {
            throw new McpException("OneNote returned unreadable image data", e);
        }

        // An image that hasn't synced can come back empty
        return data.Length > 0 ? data : throw new McpException("OneNote returned no image data. It may not be downloaded yet.");
    }

    public string FindPages(string query, string startNodeId = "") =>
        Call(app => { app.FindPages(startNodeId, query, out string xml, false, false, XMLSchema.xs2013); return xml; });

    public string GetHyperlink(string objectId) =>
        Call(app => { app.GetHyperlinkToObject(objectId, "", out string link); return link; });

    public string CreateNewPage(string sectionId) =>
        Call(app => { app.CreateNewPage(sectionId, out string pageId, NewPageStyle.npsBlankPageWithTitle); return pageId; });

    public void UpdatePageContent(string xml) =>
        Call(app => { app.UpdatePageContent(xml, DateTime.MinValue, XMLSchema.xs2013, false); return 0; });

    /// <param name="permanently">Skip the notebook's recycle bin</param>
    public void DeleteHierarchy(string objectId, bool permanently = false) =>
        Call(app => { app.DeleteHierarchy(objectId, DateTime.MinValue, permanently); return 0; });

    private T Call<T>(Func<Application2Class, T> call)
    {
        lock (_lock)
        {
            try
            {
                _app ??= new Application2Class();
                return call(_app);
            }
            catch (COMException e)
            {
                // RPC errors mean OneNote closed or restarted: reconnect on the next call.
                if ((uint)e.HResult is 0x800706BA or 0x800706BE or 0x80010108)
                {
                    _app = null;
                }

                throw new McpException(OneNoteErrors.Describe(e.HResult), e);
            }
        }
    }
}
