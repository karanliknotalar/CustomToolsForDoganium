using System.Text;
using System.Windows.Forms;
using CustomToolsForDoganium.Helper;
using CustomToolsForDoganium.Interface;
using CustomToolsForDoganium.Manager;

namespace CustomToolsForDoganium.Action
{
    internal sealed class EmptyTemplateInsertCorporateAction : IHotkeyAction
    {
        public Keys VirtualKey => Keys.T;
        public TargetApp SupportedApps => TargetApp.NotepadPlusPlus;

        public void Execute(string clipboardText, TargetApp activeApp, NotifyIcon notifyIcon)
        {
            var template = new StringBuilder()
                .AppendLine("###############################################################################")
                .AppendLine("Tel: ")
                .AppendLine()
                .AppendLine("Sigortalı: ")
                .AppendLine("Vergi: ")
                .AppendLine("Plaka: ")
                .AppendLine("Seri: ")
                .AppendLine("Motor: ")
                .AppendLine("Şasi: ")
                .AppendLine("Kullanım Şekli: ")
                .AppendLine("Model Yılı: ")
                .AppendLine("Marka Kodu: ")
                .AppendLine()
                .ToString();

            ClipboardPasteHelper.PasteAndRestore(template);
        }
    }
}