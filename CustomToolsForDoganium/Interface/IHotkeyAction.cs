using System.Windows.Forms;
using CustomToolsForDoganium.Manager;

namespace CustomToolsForDoganium.Interface
{
    internal interface IHotkeyAction
    {
        int VirtualKey { get; }
        TargetApp SupportedApps { get; }

        /// <summary>clipboardText pano boşsa null olabilir.</summary>
        void Execute(string clipboardText, TargetApp activeApp, NotifyIcon notifyIcon);
    }
}