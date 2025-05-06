using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;

[assembly: AssemblyVersion("1.0.0.6")]
[assembly: AssemblyFileVersion("1.0.0.6")]
[assembly: AssemblyInformationalVersion("v1.0.0.6")]
[assembly: AssemblyCompany("Voltura AB")]
[assembly: AssemblyConfiguration("Release")]
[assembly: AssemblyCopyright("© 2025 Voltura AB")]
[assembly: AssemblyDescription("Instantly access your five most recent folders from the system tray.")]
[assembly: AssemblyProduct("QuickFolders")]
[assembly: AssemblyTitle("QuickFolders")]
[assembly: AssemblyMetadata("RepositoryUrl", "https://github.com/voltura/QuickFolders")]

static class Program
{
    private static Config _Config;
    private static readonly Mutex _Mutex = new Mutex(true, "6E56B35E-17AB-4601-9D7C-52DE524B7A2D");
    private static readonly ToolStripRenderer DarkRenderer = new DarkMenuRenderer();
    private static readonly ToolStripRenderer DefaultRenderer = new ToolStripProfessionalRenderer();

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);

    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

    [STAThread]
    static void Main()
    {
        if (!_Mutex.WaitOne(TimeSpan.Zero, true))
        {
            return;
        }

        SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        _Config = Config.Load();

        NotifyIcon icon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables(@"%SystemRoot%\explorer.exe")),
            Visible = true,
            Text = null
        };

        DarkToolStripDropDownMenu menu = new DarkToolStripDropDownMenu
        {
            ShowCheckMargin = false,
            ShowImageMargin = true,
            ShowItemToolTips = false,
            AutoClose = true,
            AutoSize = true,
            AllowDrop = false,
            AllowItemReorder = false,
            TopLevel = true,
            Font = GetScaledMenuFont()
        };

        System.Windows.Forms.Timer mouseCheckTimer = new System.Windows.Forms.Timer
        {
            Interval = 200
        };

        System.Windows.Forms.Timer autoCloseTimer = new System.Windows.Forms.Timer
        {
            Interval = 1500
        };

        List<ToolStripDropDown> openSubMenus = new List<ToolStripDropDown>();

        mouseCheckTimer.Tick += delegate
        {
            if (!menu.Visible)
            {
                return;
            }

            Point cursorPos = Cursor.Position;

            if (menu.Bounds.Contains(cursorPos))
            {
                autoCloseTimer.Stop();

                return;
            }

            foreach (ToolStripDropDown submenu in openSubMenus)
            {
                if (submenu.Bounds.Contains(cursorPos))
                {
                    autoCloseTimer.Stop();

                    return;
                }
            }

            if (!autoCloseTimer.Enabled)
            {
                autoCloseTimer.Start();
            }
        };

        autoCloseTimer.Tick += delegate
        {
            autoCloseTimer.Stop();

            if (!menu.Visible)
            {
                return;
            }

            Point cursorPos = Cursor.Position;

            if (menu.Bounds.Contains(cursorPos))
            {
                return;
            }

            foreach (ToolStripDropDown submenu in openSubMenus)
            {
                if (submenu.Bounds.Contains(cursorPos))
                {
                    return;
                }
            }

            menu.Close();
        };

        menu.Closed += delegate
        {
            mouseCheckTimer.Stop();
            autoCloseTimer.Stop();
        };

        MouseEventHandler populator = (sender, e) =>
        {
            if (menu.Visible)
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

            DarkToolStripMenuItem header = new DarkToolStripMenuItem("QuickFolders by Voltura AB - " + Application.ProductVersion)
            {
                Image = GetThemeImage("folder"),
                Tag = "folder",
                Enabled = false
            };

            menu.Items.Add(header);
            int added = 1;

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

                    DarkToolStripMenuItem item = new DarkToolStripMenuItem(path)
                    {
                        Image = GetThemeImage(added.ToString()),
                        Tag = added.ToString()
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
                                int index;

                                if (command.StartsWith("\""))
                                {
                                    index = command.IndexOf('"', 1);
                                    executable = command.Substring(1, index - 1);
                                    arguments = command.Substring(index + 1);
                                }
                                else
                                {
                                    index = command.IndexOf(' ');
                                    executable = command.Substring(0, index);
                                    arguments = command.Substring(index + 1);
                                }

                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = executable,
                                    Arguments = arguments,
                                    UseShellExecute = false
                                });
                            }
                            else
                            {
                                Process.Start(path);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error opening folder: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };

                    menu.Items.Add(item);

                    if (added == 5)
                    {
                        break;
                    }

                    added++;
                }
                catch
                {
                }
            }

            DarkToolStripMenuItem menuRoot = new DarkToolStripMenuItem("Menu")
            {
                Image = GetThemeImage("more"),
                Tag = "more"
            };
            
            TrackSubMenu(menuRoot, openSubMenus);

            DarkToolStripMenuItem web = new DarkToolStripMenuItem("QuickFolders web site")
            {
                Image = GetThemeImage("link"),
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

            DarkToolStripMenuItem startWithWindows = new DarkToolStripMenuItem("Start with Windows")
            {
                Image = GetThemeImage("bolt"),
                Checked = StartWithWindows,
                CheckOnClick = true
            };

            DarkToolStripMenuItem folderAction = new DarkToolStripMenuItem("Folder Action...")
            {
                Image = GetThemeImage("folder"),
                Tag = "folder"
            };

            folderAction.Click += delegate
            {
                string current = _Config.FolderActionCommand ?? "";
                string input = current;

                while (true)
                {
                    input = Microsoft.VisualBasic.Interaction.InputBox("Enter command to open folder (use %1 for folder path):", "Folder Action", input);

                    if (string.IsNullOrWhiteSpace(input) || input == current)
                    {
                        return;
                    }

                    if (!input.Contains("%1"))
                    {
                        MessageBox.Show("Command must include %1 placeholder for the folder path.", "Invalid Command", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        continue;
                    }

                    input = input.Replace("\"%1\"", "%1");
                    int index = input.IndexOf("%1");

                    bool hasSpaceBefore = (index > 0 && input[index - 1] == ' ');

                    if (!hasSpaceBefore)
                    {
                        input = input.Replace("%1", " %1");
                        index++;
                    }

                    if (input == current)
                    {
                        return;
                    }

                    _Config.FolderActionCommand = input;
                    _Config.Save();

                    return;
                }
            };

            menuRoot.DropDownItems.Add(folderAction);

            DarkToolStripMenuItem setDefaultAction = new DarkToolStripMenuItem("Set Default Folder Action")
            {
                Image = GetThemeImage("exit"),
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
                setDefaultAction.Enabled = !string.IsNullOrEmpty(_Config.FolderActionCommand);
            };

            startWithWindows.CheckedChanged += (s2, e2) =>
            {
                StartWithWindows = startWithWindows.Checked;
            };

            menuRoot.DropDownItems.Add(startWithWindows);

            DarkToolStripMenuItem theme = new DarkToolStripMenuItem("Theme")
            {
                Image = GetThemeImage("theme"),
                Tag = "theme"
            };

            TrackSubMenu(theme, openSubMenus);

            DarkToolStripMenuItem systemTheme = new DarkToolStripMenuItem(Theme.System.ToString())
            {
                Image = GetThemeImage("system"),
                Tag = "system",
                Checked = _Config.AppTheme == Theme.System
            };

            DarkToolStripMenuItem darkTheme = new DarkToolStripMenuItem(Theme.Dark.ToString())
            {
                Image = GetThemeImage("darkmode"),
                Tag = "darkmode",
                Checked = _Config.AppTheme == Theme.Dark
            };

            DarkToolStripMenuItem lightTheme = new DarkToolStripMenuItem(Theme.Light.ToString())
            {
                Image = GetThemeImage("lightmode"),
                Tag = "lightmode",
                Checked = _Config.AppTheme == Theme.Light
            };

            darkTheme.Click += delegate
            {
                _Config.AppTheme = Theme.Dark;
                _Config.Save();
                ApplyTheme(menu);
                systemTheme.Checked = false;
                lightTheme.Checked = false;
                darkTheme.Checked = true;
            };

            lightTheme.Click += delegate
            {
                _Config.AppTheme = Theme.Light;
                _Config.Save();
                ApplyTheme(menu);
                systemTheme.Checked = false;
                lightTheme.Checked = true;
                darkTheme.Checked = false;
            };

            systemTheme.Click += delegate
            {
                _Config.AppTheme = Theme.System;
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


            DarkToolStripMenuItem fontSize = new DarkToolStripMenuItem("Font Size")
            {
                Image = GetThemeImage("system"),
                Tag = "system"
            };

            TrackSubMenu(fontSize, openSubMenus);

            DarkToolStripMenuItem smallFont = new DarkToolStripMenuItem(FontSize.Small.ToString())
            {
                Checked = _Config.MenuFontSize == FontSize.Small,
            };

            DarkToolStripMenuItem mediumFont = new DarkToolStripMenuItem(FontSize.Medium.ToString())
            {
                Checked = _Config.MenuFontSize == FontSize.Medium
            };

            DarkToolStripMenuItem largeFont = new DarkToolStripMenuItem(FontSize.Large.ToString())
            {
                Checked = _Config.MenuFontSize == FontSize.Large
            };

            smallFont.Click += delegate
            {
                _Config.MenuFontSize = FontSize.Small;
                _Config.Save();
                smallFont.Checked = true;
                mediumFont.Checked = false;
                largeFont.Checked = false;
            };

            mediumFont.Click += delegate
            {
                _Config.MenuFontSize = FontSize.Medium;
                _Config.Save();
                smallFont.Checked = false;
                mediumFont.Checked = true;
                largeFont.Checked = false;
            };

            largeFont.Click += delegate
            {
                _Config.MenuFontSize = FontSize.Large;
                _Config.Save();
                smallFont.Checked = false;
                mediumFont.Checked = false;
                largeFont.Checked = true;
            };

            fontSize.DropDownItems.Add(smallFont);
            fontSize.DropDownItems.Add(mediumFont);
            fontSize.DropDownItems.Add(largeFont);

            menuRoot.DropDownItems.Add(fontSize);

            DarkToolStripMenuItem exit = new DarkToolStripMenuItem("Exit")
            {
                Image = GetThemeImage("x"),
                Tag = "x"
            };

            exit.Click += delegate
            {
                icon.Visible = false;
                Application.Exit();
            };

            menuRoot.DropDownItems.Add(exit);

            menu.Items.Add(menuRoot);

            menu.Font = GetScaledMenuFont();
            ApplyTheme(menu);

            menu.Show(TaskbarMenuPositioner.GetMenuPosition(menu.Size));
            menu.Focus();

            mouseCheckTimer.Start();
        };

        icon.MouseClick += populator;
        icon.MouseMove += populator;

        Application.Run(new ApplicationContext());
    }

    private static Font GetScaledMenuFont()
    {
        float scaleFactor;

        switch (_Config.MenuFontSize)
        {
            case FontSize.Small:
                scaleFactor = 0.8f;
                break;
            case FontSize.Large:
                scaleFactor = 1.2f;
                break;
            default:
                return SystemFonts.MenuFont;
        }

        return new Font(SystemFonts.MenuFont.FontFamily, SystemFonts.MenuFont.Size * scaleFactor, SystemFonts.MenuFont.Style);
    }

    private static Image GetThemeImage(string baseName, string overrideSuffix = null)
    {
        string suffix = _Config.AppTheme == Theme.Dark ? "_dark" : "";
        string resourceName = "QuickFolders.Resources." + baseName + (overrideSuffix ?? suffix) + ".png";

        using (Stream stream = typeof(Program).Assembly.GetManifestResourceStream(resourceName))
        {
            return stream == null ? null : Image.FromStream(stream);
        }
    }

    private static void TrackSubMenu(DarkToolStripMenuItem item, List<ToolStripDropDown> submenus)
    {
        item.DropDownOpened += (s, e) =>
        {
            if (item.DropDown != null && !submenus.Contains(item.DropDown))
            {
                submenus.Add(item.DropDown);
            }
        };

        item.DropDownClosed += (s, e) =>
        {
            if (item.DropDown != null)
            {
                submenus.Remove(item.DropDown);
            }
        };
    }

    class DarkToolStripDropDownMenu : ToolStripDropDownMenu
    {
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle &= ~0x00020000; // Remove CS_DROPSHADOW
                cp.ExStyle |= 0x00000080;

                return cp;
            }
        }
    }

    class DarkToolStripMenuItem : ToolStripMenuItem
    {
        public DarkToolStripMenuItem() : base() { }

        public DarkToolStripMenuItem(string text) : base(text) { }

        protected override ToolStripDropDown CreateDefaultDropDown()
        {
            DarkToolStripDropDownMenu dropDown = new DarkToolStripDropDownMenu
            {
                OwnerItem = this
            };

            return dropDown;
        }
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
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding;
            Color textColor = (e.Item.Enabled) ? Color.FromArgb(240, 240, 240) : Color.FromArgb(160, 160, 160);
            TextRenderer.DrawText(e.Graphics, e.Text, e.TextFont, e.TextRectangle, textColor, flags);
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rect = new Rectangle(Point.Empty, e.Item.Size);

            if (e.Item.Selected)
            {
                using (System.Drawing.Drawing2D.LinearGradientBrush brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    rect,
                    Color.FromArgb(70, 70, 70),
                    Color.FromArgb(50, 50, 50),
                    System.Drawing.Drawing2D.LinearGradientMode.Vertical))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }
            }
            else if (!e.Item.Enabled)
            {
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(32, 32, 32)), rect);
            }
            else
            {
                base.OnRenderMenuItemBackground(e);
            }
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.ArrowColor = Color.Gainsboro;
            base.OnRenderArrow(e);
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

    private static void ApplyTheme(DarkToolStripDropDownMenu menu)
    {
        string suffix = _Config.AppTheme == Theme.Dark ? "_dark" : "";
        Theme realSystemTheme = _Config.AppTheme == Theme.Dark ? Theme.Dark : Theme.Light;

        if (_Config.AppTheme == Theme.System)
        {
            realSystemTheme = GetSystemTheme();
        }

        menu.Font = GetScaledMenuFont();
        menu.Renderer = realSystemTheme == Theme.Dark ? DarkRenderer : DefaultRenderer;

        foreach (ToolStripItem item in menu.Items)
        {
            ApplyThemeToMenuItem(item, suffix, realSystemTheme);
        }
    }

    public static Theme GetSystemTheme()
    {
        try
        {
            const string key = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

            RegistryKey personalizeKey = Registry.CurrentUser.OpenSubKey(key);
            if (personalizeKey != null)
            {
                object value = personalizeKey.GetValue("AppsUseLightTheme");

                if (value != null && value is int)
                {
                    int lightTheme = (int)value;

                    return lightTheme == 0 ? Theme.Dark : Theme.Light;
                }

                personalizeKey.Close();
            }
        }
        catch
        {
        }

        return Theme.Light;
    }

    private static void ApplyThemeToMenuItem(ToolStripItem item, string suffix, Theme realSystemTheme)
    {
        DarkToolStripMenuItem menuItem = item as DarkToolStripMenuItem;

        if (menuItem != null)
        {
            if (menuItem.Tag != null)
            {
                string baseName = menuItem.Tag.ToString();
                menuItem.Image = GetThemeImage(baseName, suffix);
            }

            if (realSystemTheme == Theme.Dark)
            {
                menuItem.BackColor = Color.FromArgb(32, 32, 32);
                menuItem.ForeColor = Color.Gainsboro;
            }
            else if (realSystemTheme == Theme.Light)
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
                ApplyThemeToMenuItem(subItem, suffix, realSystemTheme);
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

public enum FontSize
{
    Small,
    Medium,
    Large
}

public enum Theme
{
    System,
    Dark,
    Light
}

public class Config
{
    public string FolderActionCommand = string.Empty;

    public Theme AppTheme = Theme.System;

    public FontSize MenuFontSize = FontSize.Medium;

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
                    config.FolderActionCommand = value;
                    break;
                case "AppTheme":
                    Theme parsedTheme;
                    if (Enum.TryParse(value, out parsedTheme))
                    {
                        config.AppTheme = parsedTheme;
                    }
                    break;
                case "MenuFontSize":
                    FontSize parsedSize;
                    if (Enum.TryParse(value, out parsedSize))
                    {
                        config.MenuFontSize = parsedSize;
                    }
                    break;
            }
        }

        return config;
    }

    public void Save()
    {
        List<string> lines = new List<string>
        {
            "FolderActionCommand=" + (FolderActionCommand ?? ""),
            "AppTheme=" + AppTheme.ToString(),
            "MenuFontSize=" + MenuFontSize.ToString()
        };

        File.WriteAllLines(ConfigFilePath, lines);
    }
}

static class TaskbarMenuPositioner
{
    private enum TaskbarPosition : byte
    {
        Unknown,
        Left,
        Top,
        Right,
        Bottom
    }

    public static Point GetMenuPosition(Size menuSize)
    {
        Point position = Cursor.Position;

        TaskbarPosition taskbarPos;
        Rectangle taskbarRect;
        bool hasTaskbar = GetTaskbarPosition(out taskbarPos, out taskbarRect);

        if (hasTaskbar)
        {
            if (taskbarPos == TaskbarPosition.Bottom)
            {
                position.Y = taskbarRect.Top - menuSize.Height;
            }
            else if (taskbarPos == TaskbarPosition.Top)
            {
                position.Y = taskbarRect.Bottom;
            }
            else if (taskbarPos == TaskbarPosition.Left)
            {
                position.X = taskbarRect.Right;
            }
            else if (taskbarPos == TaskbarPosition.Right)
            {
                position.X = taskbarRect.Left - menuSize.Width;
            }
        }

        Rectangle screenBounds = Screen.FromPoint(position).WorkingArea;

        if (position.X + menuSize.Width > screenBounds.Right)
        {
            position.X = screenBounds.Right - menuSize.Width;
        }

        if (position.X < screenBounds.Left)
        {
            position.X = screenBounds.Left;
        }

        if (position.Y + menuSize.Height > screenBounds.Bottom)
        {
            position.Y = screenBounds.Bottom - menuSize.Height;
        }

        if (position.Y < screenBounds.Top)
        {
            position.Y = screenBounds.Top;
        }

        return position;
    }

    private static bool GetTaskbarPosition(out TaskbarPosition position, out Rectangle rect)
    {
        APPBARDATA data = new APPBARDATA
        {
            cbSize = Marshal.SizeOf(typeof(APPBARDATA))
        };

        IntPtr result = SHAppBarMessage(ABM_GETTASKBARPOS, ref data);

        rect = data.rc.ToRectangle();
        position = TaskbarPosition.Unknown;

        if (result == IntPtr.Zero || rect.Width == 0 || rect.Height == 0)
        {
            return false;
        }

        switch (data.uEdge)
        {
            case 0: position = TaskbarPosition.Left; break;
            case 1: position = TaskbarPosition.Top; break;
            case 2: position = TaskbarPosition.Right; break;
            case 3: position = TaskbarPosition.Bottom; break;
        }

        return true;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct APPBARDATA
    {
        public int cbSize;
        public IntPtr hWnd;
        public uint uCallbackMessage;
        public uint uEdge;
        public RECT rc;
        public IntPtr lParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int left, top, right, bottom;

        public Rectangle ToRectangle()
        {
            return Rectangle.FromLTRB(left, top, right, bottom);
        }
    }

    [DllImport("shell32.dll", CallingConvention = CallingConvention.StdCall)]
    private static extern IntPtr SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);

    private const uint ABM_GETTASKBARPOS = 5;
}
