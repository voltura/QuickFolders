using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

static class P
{
    private static readonly Mutex _Mutex = new Mutex(true, "6E56B35E-17AB-4601-9D7C-52DE524B7A2D");
    private static Config _Config;
    private static readonly ToolStripRenderer DarkRenderer = new DarkMenuRenderer();
    private static readonly ToolStripRenderer DefaultRenderer = new ToolStripProfessionalRenderer();

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
            ShowImageMargin = true,
            Font = SystemFonts.MenuFont
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
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("folder")),
                Tag = "folder",
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
                        Image = ResourceHelper.GetEmbeddedImage(GetThemeImage(menu.Items.Count.ToString())),
                        Tag = menu.Items.Count.ToString()
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
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("more")),
                Tag = "more"
            };

            ToolStripMenuItem web = new ToolStripMenuItem("QuickFolders web site")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("link")),
                Tag = "link"
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
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("bolt")),
                Checked = StartWithWindows,
                CheckOnClick = true
            };

            ToolStripMenuItem folderAction = new ToolStripMenuItem("Folder Action...")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("folder")),
                Tag = "folder"
            };

            folderAction.Click += delegate
            {
                string current = _Config.FolderActionCommand ?? "";
                string input = Microsoft.VisualBasic.Interaction.InputBox("Enter command to open folder (use %1 for folder path):", "Folder Action", current);

                if (!string.IsNullOrWhiteSpace(input) && input != current)
                {
                    _Config.FolderActionCommand = input;
                    _Config.Save();
                }
            };

            menuRoot.DropDownItems.Add(folderAction);

            ToolStripMenuItem setDefaultAction = new ToolStripMenuItem("Set Default Folder Action")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("exit")),
                Tag = "exit"
            };

            setDefaultAction.Click += delegate
            {
                if (MessageBox.Show("Reset the folder action to default (no custom command)?", "Confirm Reset", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    _Config.FolderActionCommand = null;
                    _Config.Save();
                }
            };

            menuRoot.DropDownItems.Add(setDefaultAction);

            menuRoot.DropDownOpening += (s, ee) =>
            {
                if (string.IsNullOrEmpty(_Config.FolderActionCommand))
                {
                    setDefaultAction.Enabled = false;
                }
                else
                {
                    setDefaultAction.Enabled = true;
                }
            };

            startWithWindows.CheckedChanged += (s2, e2) =>
            {
                StartWithWindows = startWithWindows.Checked;
            };

            menuRoot.DropDownItems.Add(startWithWindows);

            ToolStripMenuItem theme = new ToolStripMenuItem("Theme")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("theme")),
                Tag = "theme"
            };
            ToolStripMenuItem systemTheme = new ToolStripMenuItem("System")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("system")),
                Tag = "system",
                Checked = _Config.Theme == "System"
            };
            ToolStripMenuItem darkTheme = new ToolStripMenuItem("Dark")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("darkmode")),
                Tag = "darkmode",
                Checked = _Config.Theme == "Dark"

            };
            ToolStripMenuItem lightTheme = new ToolStripMenuItem("Light")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("lightmode")),
                Tag = "lightmode",
                Checked = _Config.Theme == "Light"
            };

            darkTheme.Click += delegate
            {
                _Config.Theme = "Dark";
                _Config.Save();
                ApplyTheme(menu);
                systemTheme.Checked = false;
                lightTheme.Checked = false;
                darkTheme.Checked = true;
            };

            lightTheme.Click += delegate
            {
                _Config.Theme = "Light";
                _Config.Save();
                ApplyTheme(menu);
                systemTheme.Checked = false;
                lightTheme.Checked = true;
                darkTheme.Checked = false;
            };

            systemTheme.Click += delegate
            {
                _Config.Theme = "System";
                _Config.Save();
                ApplyTheme(menu);
                systemTheme.Checked = true;
                lightTheme.Checked = false;
                darkTheme.Checked = false;
            };

            theme.DropDownItems.Add(systemTheme);
            theme.DropDownItems.Add(darkTheme);
            theme.DropDownItems.Add(lightTheme);

            menuRoot.DropDownItems.Add(theme);

            ToolStripMenuItem exit = new ToolStripMenuItem("Exit")
            {
                Image = ResourceHelper.GetEmbeddedImage(GetThemeImage("x")),
                Tag = "x"
            };

            exit.Click += delegate
            {
                icon.Visible = false;
                Application.Exit();
            };

            menuRoot.DropDownItems.Add(exit);

            menu.Items.Add(menuRoot);

            ApplyTheme(menu);

            menu.Show(Cursor.Position);

            menuOpen = true;

            menu.Closed += delegate
            {
                menuOpen = false;
            };
        };

        Application.Run(new ApplicationContext());
    }

    private static string GetThemeImage(string baseName)
    {
        string suffix = _Config.Theme == "Dark" ? "_dark" : "";

        return "QuickFolders.Resources." + baseName + suffix + ".png";
    }

    class DarkMenuRenderer : ToolStripProfessionalRenderer
    {
        public DarkMenuRenderer() : base(new DarkColorTable())
        {
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using (Pen p = new Pen(Color.FromArgb(40, 40, 40), 3))
            {
                Rectangle rect = new Rectangle(Point.Empty, e.ToolStrip.Size);
                rect.Width -= 1;
                rect.Height -= 1;
                e.Graphics.DrawRectangle(p, rect);
            }
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter;
            Color textColor = Color.FromArgb(240, 240, 240);
            TextRenderer.DrawText(e.Graphics, e.Text, e.TextFont, e.TextRectangle, textColor, flags);
        }
    }

    class DarkColorTable : ProfessionalColorTable
    {
        public override Color MenuItemSelected
        {
            get { return Color.FromArgb(40, 40, 40); }
        }

        public override Color MenuItemSelectedGradientBegin
        {
            get { return Color.FromArgb(40, 40, 40); }
        }

        public override Color MenuItemSelectedGradientEnd
        {
            get { return Color.FromArgb(40, 40, 40); }
        }

        public override Color MenuItemBorder
        {
            get { return Color.DarkGray; }
        }

        public override Color ToolStripBorder
        {
            get { return Color.FromArgb(40, 40, 40); }
        }

        public override Color MenuBorder
        {
            get { return Color.FromArgb(40, 40, 40); }
        }
    }

    private static void ApplyTheme(ContextMenuStrip menu)
    {
        string suffix = _Config.Theme == "Dark" ? "_dark" : "";

        menu.Renderer = (_Config.Theme == "Dark") ? DarkRenderer : DefaultRenderer;

        foreach (ToolStripItem item in menu.Items)
        {
            ApplyThemeToMenuItem(item, suffix);
        }
    }

    private static void ApplyThemeToMenuItem(ToolStripItem item, string suffix)
    {
        ToolStripMenuItem menuItem = item as ToolStripMenuItem;

        if (menuItem != null)
        {
            if (menuItem.Tag != null)
            {
                string baseName = menuItem.Tag.ToString();
                menuItem.Image = ResourceHelper.GetEmbeddedImage("QuickFolders.Resources." + baseName + suffix + ".png");
            }
            if (_Config.Theme == "Dark")
            {
                menuItem.BackColor = Color.FromArgb(32, 32, 32);
                menuItem.ForeColor = Color.Gainsboro;
            }
            else if (_Config.Theme == "Light")
            {
                menuItem.BackColor = Color.WhiteSmoke;
                menuItem.ForeColor = Color.Black;
            }
            else
            {
                menuItem.BackColor = SystemColors.Menu;
                menuItem.ForeColor = SystemColors.MenuText;
            }

            foreach (ToolStripItem subItem in menuItem.DropDownItems)
            {
                ApplyThemeToMenuItem(subItem, suffix);
            }
        }
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
    public string FolderActionCommand = string.Empty;

    public string Theme = "System";

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
        if (!File.Exists(ConfigFilePath))
        {
            Config newConfig = new Config();
            newConfig.Save();

            return newConfig;
        }

        Config config = new Config();

        string[] lines = File.ReadAllLines(ConfigFilePath);

        foreach (string line in lines)
        {
            string trimmed = line.Trim();

            if (string.IsNullOrEmpty(trimmed))
            {
                continue;
            }

            if (trimmed.StartsWith("#"))
            {
                continue;
            }

            string[] parts = trimmed.Split(new[] { '=' }, 2);

            if (parts.Length != 2)
            {
                continue;
            }

            string key = parts[0].Trim();
            string value = parts[1].Trim();

            switch (key)
            {
                case "FolderActionCommand":
                    {
                        config.FolderActionCommand = value;
                        break;
                    }
                case "Theme":
                    {
                        config.Theme = value;
                        break;
                    }
            }
        }

        if (string.IsNullOrEmpty(config.Theme))
        {
            config.Theme = "System";
        }

        return config;
    }

    public void Save()
    {
        List<string> lines = new List<string>
        {
            "FolderActionCommand=" + (FolderActionCommand ?? ""),
            "Theme=" + (Theme ?? "System")
        };

        File.WriteAllLines(ConfigFilePath, lines);
    }
}

static class ResourceHelper
{
    public static Image GetEmbeddedImage(string resourceName)
    {
        using (Stream stream = typeof(P).Assembly.GetManifestResourceStream(resourceName))
        {
            return (stream == null) ? null : Image.FromStream(stream);
        }
    }
}
