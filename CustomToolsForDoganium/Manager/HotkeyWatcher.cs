using System;

namespace CustomToolsForDoganium.Manager
{
    /// <summary>Ctrl + verilen tuşu, basılı tutmada tekrar tetiklenmeyecek şekilde (edge-triggered) izler.</summary>
    internal sealed class HotkeyWatcher(int vkKey)
    {
        private bool _wasKeyDown;

        public bool IsPressed()
        {
            var ctrlDown = (Win32Native.GetAsyncKeyState(Win32Native.VK_CONTROL) & 0x8000) != 0;
            var keyDown = (Win32Native.GetAsyncKeyState(vkKey) & 0x8000) != 0;

            var triggered = ctrlDown && keyDown && !_wasKeyDown;
            _wasKeyDown = keyDown;
            // VkScanner.ScanOnce();
            return triggered;
        }
    }
    
    internal static class VkScanner
    {
        /// <summary>Geçici debug: OEM aralığındaki tüm kodları tarar, basılı olanı loglar.</summary>
        public static void ScanOnce()
        {
            for (int vk = 0xBA; vk <= 0xE5; vk++)
            {
                if ((Win32Native.GetAsyncKeyState(vk) & 0x8000) != 0)
                {
                    Console.WriteLine($"Basılı VK: 0x{vk:X2} ({vk})");
                }
            }
        }
    }
}