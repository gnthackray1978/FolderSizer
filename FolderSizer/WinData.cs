using System;
using System.Collections.Generic;

namespace FolderSizer
{
    public struct WinData
    {
        public IntPtr Hwnd { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public List<SubWin> SubFolders { get; set; }
    }
}