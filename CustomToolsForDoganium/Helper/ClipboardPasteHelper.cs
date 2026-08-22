using System.Threading;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Helper;

internal static class ClipboardPasteHelper
{
    /// <summary>Panoyu geçici olarak verilen metinle değiştirir, yapıştırır, eski içeriği geri yükler.</summary>
    internal static void PasteAndRestore(string textToPaste)
    {
        var previousClipboard = Clipboard.ContainsText() ? Clipboard.GetText() : null;

        Clipboard.SetText(textToPaste);
        SendKeys.SendWait("^v");
        Thread.Sleep(150);

        if (previousClipboard != null)
            Clipboard.SetText(previousClipboard);
    }
}