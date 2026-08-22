using System.Collections.Generic;
using System.Windows.Forms;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.Hotkeys
{
    /// <summary>Kayıtlı tüm <see cref="IHotkeyAction"/>'ları her tick'te kontrol eder ve
    /// önplandaki uygulamayla eşleşeni tetikler. Yeni bir kısayol eklemek için sadece
    /// yeni bir IHotkeyAction implementasyonu yazıp kayıt listesine eklemek yeterlidir.</summary>
    internal sealed class HotkeyActionManager
    {
        private readonly List<(IHotkeyAction Action, HotkeyWatcher Watcher)> _entries = new();
        private readonly NotifyIcon _notifyIcon;

        internal HotkeyActionManager(NotifyIcon notifyIcon, IEnumerable<IHotkeyAction> actions)
        {
            _notifyIcon = notifyIcon;
            foreach (var action in actions)
                _entries.Add((action, new HotkeyWatcher(action.VirtualKey)));
        }

        /// <summary>Ana döngüde her tick çağrılır.</summary>
        internal void Poll()
        {
            foreach (var (action, watcher) in _entries)
            {
                if (!watcher.IsPressed()) continue;

                var processName = Win32Native.GetForegroundProcessName();
                if (processName == null) continue;

                var activeApp = ResolveTargetApp(processName);
                if (activeApp == TargetApp.None) continue;
                if ((action.SupportedApps & activeApp) == 0) continue;

                var clipboardText = Clipboard.ContainsText() ? Clipboard.GetText() : null;
                action.Execute(clipboardText, activeApp, _notifyIcon);
            }
        }

        private static TargetApp ResolveTargetApp(string processName)
        {
            if (processName.Contains("notepad++")) return TargetApp.NotepadPlusPlus;
            if (processName.Contains("excel")) return TargetApp.Excel;
            return TargetApp.None;
        }
    }
}
