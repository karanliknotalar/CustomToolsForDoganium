using System.Windows.Forms;

namespace CustomToolsForDoganium.Hotkeys;

/// <summary>Belirli bir Ctrl+Tuş kombinasyonuna ve hedef uygulama(lar)a bağlı bir kısayol eylemi.</summary>
internal interface IHotkeyAction
{
    Keys VirtualKey { get; }
    TargetApp SupportedApps { get; }

    /// <summary>clipboardText pano boşsa null olabilir.</summary>
    void Execute(string clipboardText, TargetApp activeApp, NotifyIcon notifyIcon);
}