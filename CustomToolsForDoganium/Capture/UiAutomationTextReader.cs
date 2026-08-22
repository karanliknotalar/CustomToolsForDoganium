using System;
using System.Text;
using System.Windows.Automation;

namespace CustomToolsForDoganium.Capture
{
    /// <summary>Bir pencerenin UI Automation ağacını gezerek görünür/erişilebilir tüm metni toplar.</summary>
    internal static class UiAutomationTextReader
    {
        private const int MaxDepth = 7;

        internal static string ReadText(IntPtr hwnd)
        {
            var root = AutomationElement.FromHandle(hwnd);
            if (root == null) return null;

            var sb = new StringBuilder();
            Walk(root, sb, 0);
            return sb.ToString();
        }

        private static void Walk(AutomationElement element, StringBuilder sb, int depth)
        {
            if (depth > MaxDepth) return;

            if (!string.IsNullOrWhiteSpace(element.Current.Name))
                sb.AppendLine(element.Current.Name);

            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
                sb.AppendLine(((ValuePattern)valuePattern).Current.Value);

            if (element.TryGetCurrentPattern(TextPattern.Pattern, out var textPattern))
                sb.AppendLine(((TextPattern)textPattern).DocumentRange.GetText(-1));

            var children = element.FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement child in children)
                Walk(child, sb, depth + 1);
        }
    }
}
