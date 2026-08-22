using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CustomToolsForDoganium.Capture;
using CustomToolsForDoganium.Hotkeys;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.App;

/// <summary>
/// Uygulamanın çalışma zamanını kurar ve başlatır. Eski 100ms'lik busy-poll while döngüsü
/// yerine: görünmez bir pencerenin sağladığı standart WinForms mesaj pompası + bu pompaya
/// bağlı bir WH_KEYBOARD_LL klavye kancası kullanılır. Sonuç: kısayollar anında (poll
/// gecikmesi olmadan) tetiklenir ve boşta CPU kullanımı pratikte sıfıra iner.
/// </summary>
internal sealed class ApplicationRunner(
    IntPtr consoleWindow,
    NotifyIcon notifyIcon,
    IEnumerable<IHotkeyAction> actions,
    DoganiumCaptureService captureService)
{
    internal void Run()
    {
        using var messageLoopForm = new HiddenMessageLoopForm(consoleWindow);

        // .Handle erişimi, WinForms'da handle'ı henüz yoksa oluşturulmasını garanti eder
        // (bkz. Control.Handle belgeleri). Bunu kancayı kurmadan ÖNCE yapıyoruz; aksi halde
        // kullanıcı çok erken bir kısayola basarsa dispatcher'ın BeginInvoke çağrısı henüz
        // var olmayan bir handle üzerinde exception fırlatabilirdi.
        _ = messageLoopForm.Handle;

        var dispatcher = new GlobalHotkeyDispatcher(
            actions, CaptureHotkeyBindings.All, captureService, notifyIcon, messageLoopForm);

        using var keyboardHook = new KeyboardHook();
        keyboardHook.KeyDown += dispatcher.OnKeyDown;
        keyboardHook.Install();

        Application.Run(messageLoopForm);
    }
}