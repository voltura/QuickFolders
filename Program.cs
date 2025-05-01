using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

static class P
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        NotifyIcon icon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables(@"%SystemRoot%\explorer.exe")),
            Visible = true
        };

        ContextMenuStrip menu = new ContextMenuStrip();
        icon.ContextMenuStrip = menu;
        bool menuOpen = false;

        icon.MouseClick += (s, e) =>
        {
            if (menuOpen || e.Button != MouseButtons.Right)
            {
                return;
            }

            FileInfo[] files;

            try
            {
                files = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Recent)).GetFiles("*.lnk");
            }
            catch
            {
                return;
            }

            Array.Sort(files, (a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));
            menu.Items.Clear();
            ToolStripMenuItem header = new ToolStripMenuItem("QuickFolders by Voltura AB")
            {
                Enabled = false
            };
            menu.Items.Add(header);

            foreach (FileInfo file in files)
            {
                try
                {
                    IShellLink link = (IShellLink)new ShellLink();
                    ((IPersistFile)link).Load(file.FullName, 0);
                    StringBuilder target = new StringBuilder(260);
                    link.GetPath(target, target.Capacity, IntPtr.Zero, 0);
                    string path = target.ToString();

                    if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || !Directory.Exists(path))
                    {
                        continue;
                    }

                    ToolStripMenuItem item = new ToolStripMenuItem(path);
                    item.Click += delegate
                    {
                        try
                        {
                            Process.Start(path);
                        }
                        catch
                        {
                        }
                    };

                    menu.Items.Add(item);

                    if (menu.Items.Count == 5)
                    {
                        break;
                    }
                }
                catch
                {
                }
            }

            ToolStripMenuItem web = new ToolStripMenuItem("QuickFolders web site");
            web.Click += delegate
            {
                try
                {
                    Process.Start("https://voltura.github.io/QuickFolders");
                }
                catch
                {
                }
            };
            menu.Items.Add(web);

            ToolStripMenuItem close = new ToolStripMenuItem("Close");
            close.Click += delegate
            {
                menu.Close();
            };
            menu.Items.Add(close);

            ToolStripMenuItem exit = new ToolStripMenuItem("Exit");
            exit.Click += delegate
            {
                icon.Visible = false;
                Application.Exit();
            };
            menu.Items.Add(exit);

            menu.Show(Cursor.Position);
            menuOpen = true;
            menu.Closed += delegate { menuOpen = false; };
        };

        Application.Run(new ApplicationContext());
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink { }

    [ComImport]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellLink
    {
        void GetPath([Out] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out] StringBuilder pszName, int cchMaxName);
        void SetDescription(string pszName);
        void GetWorkingDirectory([Out] StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory(string pszDir);
        void GetArguments([Out] StringBuilder pszArgs, int cchMaxPath);
        void SetArguments(string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation(string pszIconPath, int iIcon);
        void SetRelativePath(string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath(string pszFile);
    }

    [ComImport]
    [Guid("0000010b-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistFile
    {
        void GetClassID(out Guid pClassID);
        void IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
    }
}
