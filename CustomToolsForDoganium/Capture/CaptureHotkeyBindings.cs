using System.Windows.Forms;

namespace CustomToolsForDoganium.Capture;

/// <summary>Ctrl+Shift+Tuş yakalama kısayollarının hangi sorgu tipine karşılık geldiğini tanımlar.</summary>
internal static class CaptureHotkeyBindings
{
    internal static readonly (Keys Key, InsuranceQueryType QueryType)[] All =
    [
        (Keys.C, InsuranceQueryType.Trafik),
        (Keys.X, InsuranceQueryType.Kasko),
        (Keys.Z, InsuranceQueryType.Imm)
    ];
}