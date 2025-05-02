using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

static class P
{
    private static readonly Mutex _Mutex = new Mutex(true, "6E56B35E-17AB-4601-9D7C-52DE524B7A2D");
    static Config _Config;

    [STAThread]
    static void Main()
    {
        if (!_Mutex.WaitOne(TimeSpan.Zero, true))
        {
            return;
        }

        _Config = Config.Load();

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        NotifyIcon icon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables(@"%SystemRoot%\explorer.exe")),
            Visible = true
        };

        ContextMenuStrip menu = new ContextMenuStrip
        {
            ShowCheckMargin = false,
            ShowImageMargin = true
        };

        menu.PreviewKeyDown += Menu_PreviewKeyDown;
        menu.KeyDown += Menu_KeyDown;

        icon.ContextMenuStrip = menu;

        bool menuOpen = false;

        icon.MouseClick += (sender, e) =>
        {
            if (menuOpen)
            {
                return;
            }

            if (e.Button != MouseButtons.Right)
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
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.folder.png"),
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

                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    if (!Path.IsPathRooted(path))
                    {
                        continue;
                    }

                    if (!Directory.Exists(path))
                    {
                        continue;
                    }

                    ToolStripMenuItem item = new ToolStripMenuItem(path)
                    {
                        Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources." + menu.Items.Count + ".png")
                    };

                    item.Click += delegate
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(_Config.FolderActionCommand))
                            {
                                string command = _Config.FolderActionCommand.Replace("%1", "\"" + path + "\"");

                                string executable;
                                string arguments;

                                if (command.StartsWith("\""))
                                {
                                    int endQuoteIndex = command.IndexOf('"', 1);

                                    executable = command.Substring(1, endQuoteIndex - 1);
                                    arguments = command.Substring(endQuoteIndex + 1).TrimStart();
                                }
                                else
                                {
                                    int spaceIndex = command.IndexOf(' ');

                                    if (spaceIndex == -1)
                                    {
                                        executable = command;
                                        arguments = "";
                                    }
                                    else
                                    {
                                        executable = command.Substring(0, spaceIndex);
                                        arguments = command.Substring(spaceIndex + 1).TrimStart();
                                    }
                                }

                                ProcessStartInfo startInfo = new ProcessStartInfo
                                {
                                    FileName = executable,
                                    Arguments = arguments,
                                    UseShellExecute = false
                                };

                                Process.Start(startInfo);
                            }
                            else
                            {
                                Process.Start(path);
                            }
                        }
                        catch
                        {
                        }
                    };

                    menu.Items.Add(item);

                    if (menu.Items.Count == 6)
                    {
                        break;
                    }
                }
                catch
                {
                }
            }

            ToolStripMenuItem menuRoot = new ToolStripMenuItem("Menu")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.more.png")
            };

            ToolStripMenuItem web = new ToolStripMenuItem("QuickFolders web site")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.link.png")
            };

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

            menuRoot.DropDownItems.Add(web);

            ToolStripMenuItem startWithWindows = new ToolStripMenuItem("Start with Windows")
            {
                Checked = StartWithWindows,
                CheckOnClick = true
            };

            ToolStripMenuItem folderAction = new ToolStripMenuItem("Folder Action...") {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.folder.png")
            };

            folderAction.Click += delegate
            {
                string current = _Config.FolderActionCommand ?? "";

                string input = Microsoft.VisualBasic.Interaction.InputBox("Enter command to open folder (use %1 for folder path):", "Folder Action", current);

                if (input != null)
                {
                    _Config.FolderActionCommand = input;
                    _Config.Save();
                }
            };

            menuRoot.DropDownItems.Add(folderAction);


            startWithWindows.CheckedChanged += (s2, e2) =>
            {
                StartWithWindows = startWithWindows.Checked;
            };

            menuRoot.DropDownItems.Add(startWithWindows);

            ToolStripMenuItem theme = new ToolStripMenuItem("Theme")
            {
                Enabled = false,
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.theme.png")
            };
            ToolStripMenuItem systemTheme = new ToolStripMenuItem("System") {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.system.png")
            };
            ToolStripMenuItem darkTheme = new ToolStripMenuItem("Dark")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.darkmode.png")
            };
            ToolStripMenuItem lightTheme = new ToolStripMenuItem("Light")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.lightmode.png")
            };

            theme.DropDownItems.Add(systemTheme);
            theme.DropDownItems.Add(darkTheme);
            theme.DropDownItems.Add(lightTheme);

            menuRoot.DropDownItems.Add(theme);
            ToolStripMenuItem close = new ToolStripMenuItem("Close")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.exit.png")
            };

            close.Click += delegate
            {
                menu.Close();
            };

            menuRoot.DropDownItems.Add(close);

            ToolStripMenuItem exit = new ToolStripMenuItem("Exit")
            {
                Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources.x.png")
            };

            exit.Click += delegate
            {
                icon.Visible = false;
                Application.Exit();
            };

            menuRoot.DropDownItems.Add(exit);

            menu.Items.Add(menuRoot);

            menu.Show(Cursor.Position);

            menuOpen = true;

            menu.Closed += delegate
            {
                menuOpen = false;
            };
        };

        Application.Run(new ApplicationContext());
    }

    private static void Menu_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
    {
        if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
        {
            e.IsInputKey = true;
        }
    }

    private static void Menu_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
        {
            int index = e.KeyCode - Keys.D1 + 1;

            ContextMenuStrip menu = sender as ContextMenuStrip;

            if (menu != null && index < menu.Items.Count)
            {
                ToolStripItem item = menu.Items[index];

                if (item != null && item.Enabled)
                {
                    item.PerformClick();
                    e.Handled = true;
                }
            }
        }
    }

    public static bool StartWithWindows
    {
        get
        {
            bool startWithWindows = false;

            try
            {
                startWithWindows = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\", Application.ProductName, null) != null;
            }
            catch
            {
            }

            return startWithWindows;
        }
        set
        {
            try
            {
                using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (value)
                    {
                        registryKey.SetValue(Application.ProductName, "\"" + Application.ExecutablePath + "\"");
                    }
                    else
                    {
                        if (registryKey.GetValue(Application.ProductName) != null)
                        {
                            registryKey.DeleteValue(Application.ProductName);
                        }
                    }
                }
            }
            catch
            {
            }
        }
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink
    {
    }

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

class Config
{
    public string FolderActionCommand;

    public static string ConfigFilePath
    {
        get
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "QuickFolders");
            Directory.CreateDirectory(folder);

            return Path.Combine(folder, "QuickFolders.config");
        }
    }

    public static Config Load()
    {
        Config config = new Config();

        if (File.Exists(ConfigFilePath))
        {
            string[] lines = File.ReadAllLines(ConfigFilePath);

            foreach (string line in lines)
            {
                if (line.StartsWith("FolderActionCommand="))
                {
                    config.FolderActionCommand = line.Substring("FolderActionCommand=".Length);
                }
            }
        }

        return config;
    }

    public void Save()
    {
        File.WriteAllText(ConfigFilePath, "FolderActionCommand=" + (FolderActionCommand ?? "") + "\r\n");
    }
}

static class ResourceHelper
{
    public static Image GetEmbeddedImage(string resourceName)
    {
        using (Stream stream = typeof(P).Assembly.GetManifestResourceStream(resourceName))
        {
            if (stream == null)
            {
                return null;
            }

            return Image.FromStream(stream);
        }
    }
}
