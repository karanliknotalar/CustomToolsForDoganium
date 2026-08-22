using System.Windows.Forms;
using CustomToolsForDoganium.Native;

namespace CustomToolsForDoganium.Capture
{
    /// <summary>Doğanium sorgu ekranı yakalama kısayollarını (Ctrl+Shift+C/X/Z) izler.
    /// Basılı tutmada tekrar tetiklenmez (edge-triggered), diğer kısayollarla tutarlıdır.</summary>
    internal sealed class CaptureHotkeyWatcher
    {
        private static readonly (Keys Key, InsuranceQueryType QueryType)[] Bindings =
        [
            (Keys.C, InsuranceQueryType.Trafik),
            (Keys.X, InsuranceQueryType.Kasko),
            (Keys.Z, InsuranceQueryType.Imm)
        ];

        private bool _wasKeyDown;

        internal bool TryGetPressedQueryType(out InsuranceQueryType queryType)
        {
            queryType = InsuranceQueryType.None;

            var ctrlShiftDown = Win32Native.IsKeyDown(Keys.ControlKey) && Win32Native.IsKeyDown(Keys.ShiftKey);
            if (!ctrlShiftDown)
            {
                _wasKeyDown = false;
                return false;
            }

            foreach (var (key, type) in Bindings)
            {
                if (!Win32Native.IsKeyDown(key)) continue;

                if (_wasKeyDown) return false; // zaten basılıydı, tekrar tetikleme
                _wasKeyDown = true;
                queryType = type;
                return true;
            }

            return false;
        }
    }
}
