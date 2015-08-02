using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace FolderSizer
{
    public static class FolderData
    {
        public static String BytesToString(long byteCount)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }

        public static int GetZOrder(IntPtr hWnd)
        {
            var z = 0;
            for (IntPtr h = hWnd; h != IntPtr.Zero; h = GetWindow(h, GetWindow_Cmd.GW_HWNDPREV)) z++;
            return z;
        }

        public static string GetExplorerWindowPath(IntPtr MyHwnd)
        {
            //  IntPtr MyHwnd = FindWindow(null, "Directory");
            var t = Type.GetTypeFromProgID("Shell.Application");
            dynamic o = Activator.CreateInstance(t);
            try
            {
                var ws = o.Windows();
                for (int i = 0; i < ws.Count; i++)
                {
                    var ie = ws.Item(i);
                    if (ie == null || ie.hwnd != (long)MyHwnd) continue;
                    var path = System.IO.Path.GetFileName((string)ie.FullName);
                    if (path.ToLower() == "explorer.exe")
                    {
                        var explorepath = ie.document.focuseditem.path;
                        return explorepath;
                    }
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(o);
            } 

            return "";
        }

        public static string GetProcessPath(IntPtr hwnd)
        {
            uint pid = 0;
            GetWindowThreadProcessId(hwnd, out pid);
            if (hwnd == IntPtr.Zero) return "";
            if (pid == 0) return "";
            var process = Process.GetProcessById((int)pid);
            try
            {
                return process.MainModule.FileName.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "";
            }
        }  

        /// <summary>Returns a dictionary that contains the handle and title of all the open windows.</summary>
        /// <returns>A dictionary that contains the handle and title of all the open windows.</returns>
        public static List<WinData> GetOpenWindows(bool isFolderOnly =true)
        {
            IntPtr shellWindow = GetShellWindow();
            var windows = new List<WinData>();

            EnumWindows(delegate(IntPtr hWnd, int lParam)
            {
                if (hWnd == shellWindow) return true;
                if (!IsWindowVisible(hWnd)) return true;

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return true;

                var builder = new StringBuilder(length);
                GetWindowText(hWnd, builder, length + 1);

                var process = GetProcessPath(hWnd);

                if (process.ToLower().Contains("explorer"))
                {
                    var subWindows = new List<SubWin>();

                    var p = GetExplorerWindowPath(hWnd);

                    var pp = Directory.GetParent(p);
                   
                    foreach (var d in Directory.GetDirectories(pp.FullName))
                    {

                        Int64 dirSize = 0;
                        var di = new DirectoryInfo(d);

                        try
                        {
                            var fso = new Scripting.FileSystemObject();
                            var folder = fso.GetFolder(d);                            
                            dirSize = (Int64)folder.Size;
                        }
                        catch (Exception e)
                        {
                            if (e.Message.ToLower().Contains("permission"))
                            {
                                Console.WriteLine(di.Name + " Permission Denied - try running as administrator");   
                            }
                            else
                            {
                                Console.WriteLine(di.Name + " Error: " + e.Message);                                   
                            }
                            
                        }
                        
                        
                        subWindows.Add(new SubWin()
                        {
                            FullName = d,
                            size = dirSize,
                            FormattedSize = BytesToString(dirSize),
                            Name = di.Name
                        });
                    }

                    windows.Add(new WinData()
                    {
                        Hwnd = hWnd,
                        Name = builder.ToString(),
                        Path = GetExplorerWindowPath(hWnd),
                        SubFolders = subWindows
                    });
                }
                return true;

            }, 0);

            return windows;
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, int lParam);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        private static extern IntPtr GetShellWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr handle, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetWindow(IntPtr hWnd, GetWindow_Cmd uCmd);

        enum GetWindow_Cmd : uint
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }
    }
}