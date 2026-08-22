using System;
using System.Windows.Forms;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.App;

/// <summary>
/// Görünmez bir pencere: <c>Application.Run(this)</c> ile başlatılan mesaj pompası hem
/// <see cref="KeyboardHook"/>'un ihtiyaç duyduğu döngüyü sağlar hem de hook thread'inden
/// gelen işlerin <c>BeginInvoke</c> ile güvenle çalıştırılacağı UI thread'ini oluşturur.
/// Asla <c>Show()</c> edilmez; <see cref="SetVisibleCore"/> override'ı bunu garanti eder.
/// Eski 100ms'lik while-loop'taki konsol küçültme kontrolü burada bir <see cref="Timer"/>'a taşındı.
/// </summary>
internal sealed class HiddenMessageLoopForm : Form
{
    private readonly Timer _consoleWatchTimer;

    internal HiddenMessageLoopForm(IntPtr consoleWindow)
    {
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.FixedToolWindow;

        _consoleWatchTimer = new Timer { Interval = 300 };
        _consoleWatchTimer.Tick += (_, _) =>
        {
            if (ConsoleWindowNative.IsMinimized(consoleWindow))
                ConsoleWindowNative.Hide(consoleWindow);
        };
        _consoleWatchTimer.Start();
    }

    protected override void SetVisibleCore(bool value) => base.SetVisibleCore(false);

    protected override void Dispose(bool disposing)
    {
        if (disposing) _consoleWatchTimer.Dispose();
        base.Dispose(disposing);
    }
}