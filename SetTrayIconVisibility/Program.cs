//
// C# (Windows 8+)
// Set Tray Icon Always Visible
// https://github.com/dkxce/SetTrayIconVisibility
// en,ru,1251,utf-8
//

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Management;
using Microsoft.Win32;

namespace TestTestBVSC
{
    public static class SetTrayIconVisibility
    {
        #region WinAPI

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, [MarshalAs(UnmanagedType.U4)] uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        #endregion WinAPI

        public enum VisibilityType : byte { Default = 0, Invisible = 1, Visible = 2 }
        
        private static string ROT13(string input)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in input)
            {
                if (char.IsLetter(c))
                {
                    if (Convert.ToInt32((char.ToUpper(c))) < 78)
                        sb.Append(Convert.ToChar(Convert.ToInt32(c) + 13));
                    else
                        sb.Append(Convert.ToChar(Convert.ToInt32(c) - 13));
                }
                else sb.Append(c);
                if (sb.Length > 2 && sb[sb.Length - 1] == (char)0 && sb[sb.Length - 2] == (char)0 && sb[sb.Length - 3] == (char)0) break;
            };
            return sb.ToString().TrimEnd(new char[] { (char)0 });
        }

        public static void SetVisibility(string appExeName /* ex: traymond2.exe */, VisibilityType vType = VisibilityType.Visible, bool reloadExplorer = true)
        {
            bool update = false;

            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\TrayNotify", true);
            byte[] ba = (byte[])rk.GetValue("IconStreams");
            int recordCount = BitConverter.ToInt32(ba, 12); // IconStreams header, 20 bytes
            for (int i = 0; i < recordCount; i++) // Each IconStreams record after the 20 byte header is 1640 bytes.
            {
                int offset = 20 + i * 1640;
                byte[] exeBin = new byte[528];
                Array.Copy(ba, offset, exeBin, 0, 528);
                string exeROT = Encoding.Unicode.GetString(exeBin);
                string exePath = ROT13(exeROT);                

                if (exePath.EndsWith(appExeName))
                {
                    ba[offset + 528] = (byte)vType;
                    update = true;
                };

                // Current Visibility
                int visibility = BitConverter.ToInt32(ba, offset + 528);
                Console.WriteLine($"{visibility} {exePath}");
            };
            if(update) rk.SetValue("IconStreams", ba);
            rk.Close();

            if (update && reloadExplorer)
            {
                try
                {
                    IntPtr ptr = FindWindow("Shell_TrayWnd", null);
                    PostMessage(ptr, 0x0400 + 436, (IntPtr)0, (IntPtr)0);

                    do
                    {
                        ptr = FindWindow("Shell_TrayWnd", null);
                        if (ptr.ToInt32() == 0) break;
                        Thread.Sleep(500);
                    } while (true);
                }
                catch { };
                Process process = new Process();
                process.StartInfo.FileName = string.Format("{0}\\{1}", Environment.GetEnvironmentVariable("WINDIR"), "explorer.exe");
                process.StartInfo.UseShellExecute = true;
                process.Start();
            };

            try // Doesn't works correctrly
            {
                IntPtr ptr = Marshal.StringToHGlobalUni("shell:::{05d7b0f4-2121-4eff-bf6b-ed3f69b894d9}");
                SHChangeNotify(0x00001000, 0x0005, ptr, IntPtr.Zero);
                Marshal.FreeHGlobal(ptr);
                SHChangeNotify(0x08000000 /* SHCNE_ASSOCCHANGED */, 0 /* SHCNF_IDLIST */, IntPtr.Zero, IntPtr.Zero);
                SendMessage((IntPtr)0xFFFF/* HWND_BROADCAST */, 0x001A /* WM_SETTINGCHANGE */, IntPtr.Zero, IntPtr.Zero);
            }
            catch { };            
        }

        private static string GetCommandLine(this Process process)
        {
            string query = $@"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}";
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
                using (ManagementObjectCollection collection = searcher.Get())
                    foreach (ManagementObject mo in collection)
                        return  (string)mo["CommandLine"];
            return null;
        }

        public static Process RestartProcess(string appExeName /* ex: traymond2.exe */)
        {
            foreach(Process proc in Process.GetProcesses())
            {
                try
                {
                    if (proc.MainModule.FileName.EndsWith(appExeName))
                    {
                        string fName = proc.MainModule.FileName;
                        string cLine = GetCommandLine(proc);
                        cLine = cLine.Replace($"\"{fName}\"", "").Replace(fName, "").Trim();
                        proc.Kill();
                        return Process.Start(fName, cLine);
                    };
                }
                catch { };
            };
            return null;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("SetTrayIconVisibility by dkxce \r\nhttps://github.com/dkxce/SetTrayIconVisibility\r\n");
                Console.WriteLine("Usage: SetTrayIconVisibility.exe -{VisibilityType} [-noRestartExplorer] [-restartApp] {appExe}");
                Console.WriteLine("Visible:   SetTrayIconVisibility.exe -2 {appExe}");
                Console.WriteLine("Invisible: SetTrayIconVisibility.exe -1 {appExe}");
                Console.WriteLine("HiddenL    SetTrayIconVisibility.exe -0 {appExe}");
                System.Threading.Thread.Sleep(2000);
                return;  
            };
            string exe = null;
            SetTrayIconVisibility.VisibilityType vType = SetTrayIconVisibility.VisibilityType.Default;
            bool noRestartExplorer = false;
            bool restartApp = false;
            foreach (string arg in args)
            {
                if (!arg.StartsWith("-") && !arg.StartsWith("/")) exe = arg;
                if (arg == "/0" || arg == "-0") vType = SetTrayIconVisibility.VisibilityType.Default;
                if (arg == "/1" || arg == "-2") vType = SetTrayIconVisibility.VisibilityType.Invisible;
                if (arg == "/1" || arg == "-2") vType = SetTrayIconVisibility.VisibilityType.Visible;
                if (arg == "/noRestartExplorer" || arg == "-noRestartExplorer") noRestartExplorer = true;
                if (arg == "/restartApp" || arg == "-restartApp") restartApp = true;
            };
            SetTrayIconVisibility.SetVisibility(exe, vType, !noRestartExplorer);
            if(restartApp) SetTrayIconVisibility.RestartProcess(exe);
        }
    }
}