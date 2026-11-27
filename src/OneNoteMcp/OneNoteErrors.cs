namespace OneNoteMcp;

/// <summary>
/// Turns OneNote COM error codes into messages an agent can act on.
/// </summary>
public static class OneNoteErrors
{
    private static readonly Dictionary<uint, string> Messages = new()
    {
        [0x80042000] = "The XML is malformed",
        [0x80042001] = "The XML is invalid for the OneNote schema",
        [0x80042004] = "Section not found",
        [0x80042005] = "Page not found",
        [0x80042006] = "File not found",
        [0x80042007] = "Couldn't insert the image",
        [0x80042009] = "Couldn't insert the HTML. OneNote text elements take inline HTML only.",
        [0x8004200B] = "The section is read-only",
        [0x8004200C] = "The page is read-only",
        [0x8004200E] = "Page object not found",
        [0x8004200F] = "OneNote hasn't downloaded this image yet. Open the page in OneNote desktop, then retry.",
        [0x80042010] = "The page changed since it was read",
        [0x80042011] = "Section group not found",
        [0x80042014] = "Object not found",
        [0x80042015] = "Notebook not found",
        [0x80042017] = "Invalid name",
        [0x80042019] = "Invalid search query",
        [0x8004201B] = "The section is password-protected and locked",
        [0x8004201C] = "Blocked by policy",
        [0x8004201D] = "OneNote hasn't synced this content yet. Open it in OneNote desktop, then retry.",
        [0x800706BA] = "OneNote isn't responding (RPC server unavailable). Retry.",
        [0x800706BE] = "OneNote's call failed (remote procedure call failed). Retry.",
        [0x80010108] = "OneNote closed during the call. Retry.",
    };

    public static string Describe(int hresult)
    {
        var code = (uint)hresult;
        return Messages.TryGetValue(code, out var message)
            ? $"{message} (0x{code:X8})"
            : $"OneNote COM error 0x{code:X8}";
    }
}
