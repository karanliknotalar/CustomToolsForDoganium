using System.Windows.Forms;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.Hotkeys
{
    /// <summary>Ctrl + verilen tuşu, basılı tutmada tekrar tetiklenmeyecek şekilde (edge-triggered) izler.</summary>
    internal sealed class HotkeyWatcher(Keys vKey)
    {
        private bool _wasKeyDown;

        internal bool IsPressed()
        {
            var ctrlDown = Win32Native.IsKeyDown(Keys.ControlKey);
            var keyDown = Win32Native.IsKeyDown(vKey);

            var triggered = ctrlDown && keyDown && !_wasKeyDown;
            _wasKeyDown = keyDown;
            return triggered;
        }
    }
}
