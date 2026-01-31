using System;
using System.Windows.Forms;

namespace CustomToolsForDoganium
{
    public static class Tools
    {
        private const string Title = "";

        public static void ShowNotifyInfo(NotifyIcon n, string m)
        {
            n.ShowBalloonTip(1000, Title,
                m, ToolTipIcon.Info);
            Console.WriteLine(m);
        }

        public static void ShowNotifyWarning(NotifyIcon n, string m)
        {
            n.ShowBalloonTip(1000, Title,
                m, ToolTipIcon.Warning);
            Console.WriteLine(m);
        }
    }
}