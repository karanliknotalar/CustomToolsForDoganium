using System;
using System.Windows.Forms;

namespace CustomToolsForDoganium.Manager
{
    /// <summary>Ctrl + verilen tuşu, basılı tutmada tekrar tetiklenmeyecek şekilde (edge-triggered) izler.</summary>
    internal sealed class HotkeyWatcher(Keys vKey)
    {
        private bool _wasKeyDown;

        public bool IsPressed()
        {
            var ctrlDown = (Win32Native.GetAsyncKeyState(Keys.ControlKey) & 0x8000) != 0;
            var keyDown = (Win32Native.GetAsyncKeyState(vKey) & 0x8000) != 0;

            var triggered = ctrlDown && keyDown && !_wasKeyDown;
            _wasKeyDown = keyDown;
            return triggered;
        }
    }
}