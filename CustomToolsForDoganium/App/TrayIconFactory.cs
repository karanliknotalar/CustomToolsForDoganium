using System;
using System.Drawing;
using System.Windows.Forms;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.App;

/// <summary>Sistem tepsisi (tray) ikonunu ve sağ tık menüsünü oluşturur.</summary>
internal static class TrayIconFactory
{
    internal static NotifyIcon Create(IntPtr consoleWindow)
    {
        var notifyIcon = new NotifyIcon
        {
            Icon = SystemIcons.Shield,
            Text = "Doganium Araçları",
            Visible = true
        };

        notifyIcon.DoubleClick += (_, _) => ConsoleWindowNative.Restore(consoleWindow);

        notifyIcon.ContextMenu = new ContextMenu([
            new MenuItem("Göster", (_, _) => ConsoleWindowNative.Restore(consoleWindow)),
            new MenuItem("-"),
            new MenuItem("Çıkış", (_, _) =>
            {
                notifyIcon.Visible = false;
                Application.Exit();
            })
        ]);

        return notifyIcon;
    }
}