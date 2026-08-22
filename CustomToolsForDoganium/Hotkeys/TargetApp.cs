using System;

namespace CustomToolsForDoganium.Hotkeys;

[Flags]
internal enum TargetApp
{
    None = 0,
    NotepadPlusPlus = 1 << 0,
    Excel = 1 << 1,
    // Yeni uygulama desteği buraya eklenir
}