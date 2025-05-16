using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

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
    private static readonly Mutex _Mutex = new Mutex(true, "6E56B35E-17AB-4601-9D7C-52DE524B7A2D");

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);

    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

    private static DarkToolStripDropDownMenu menu;
    private static System.Windows.Forms.Timer mouseCheckTimer;
    private static System.Windows.Forms.Timer autoCloseTimer;
    private static DarkToolStripMenuItem setDefaultAction;
    private static DarkToolStripMenuItem startWithWindows;
    private static DarkToolStripMenuItem systemTheme;
    private static DarkToolStripMenuItem lightTheme;
    private static DarkToolStripMenuItem darkTheme;
    private static DarkToolStripMenuItem smallFont;
    private static DarkToolStripMenuItem mediumFont;
    private static DarkToolStripMenuItem largeFont;
    private static NotifyIcon icon;
    private static bool isShowingMenu = false;
    private static readonly List<ToolStripDropDown> openSubMenus = new List<ToolStripDropDown>();


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
        ThemeHelpers.Config = Config.Load();

        icon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables(@"%SystemRoot%\explorer.exe")),
            Visible = true,
            Text = null
        };

        menu = new DarkToolStripDropDownMenu
        {
            ShowCheckMargin = false,
            ShowImageMargin = true,
            ShowItemToolTips = false,
            AutoClose = true,
            AutoSize = true,
            AllowDrop = false,
            AllowItemReorder = false,
            TopLevel = true,
            Font = ThemeHelpers.GetScaledMenuFont()
        };

        mouseCheckTimer = new System.Windows.Forms.Timer
        {
            Interval = 200
        };

        autoCloseTimer = new System.Windows.Forms.Timer
        {
            Interval = 1500
        };

        mouseCheckTimer.Tick += OnMouseCheckTimerTick;
        autoCloseTimer.Tick += OnAutoCloseTimerTick;
        menu.Closed += OnMenuClosed;
        icon.MouseClick += OnIconMouseInteraction;
        icon.MouseMove += OnIconMouseInteraction;

        Application.Run(new ApplicationContext());
    }

    private static void OnIconMouseInteraction(object sender, MouseEventArgs e)
    {
        if (isShowingMenu)
        {
            return;
        }

        isShowingMenu = true;

        if (menu != null)
        {
            menu.Closed -= OnMenuClosed;
            DisposeMenuItems(menu.Items);
            menu.Dispose();
            menu = null;
        }

        menu = new DarkToolStripDropDownMenu
        {
            ShowCheckMargin = false,
            ShowImageMargin = true,
            ShowItemToolTips = false,
            AutoClose = true,
            AutoSize = true,
            AllowDrop = false,
            AllowItemReorder = false,
            TopLevel = true,
            Font = ThemeHelpers.GetScaledMenuFont()
        };

        menu.Closed += OnMenuClosed;

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

        Array.Sort(files, delegate (FileInfo a, FileInfo b)
        {
            return b.LastWriteTime.CompareTo(a.LastWriteTime);
        });

        DarkToolStripMenuItem header = new DarkToolStripMenuItem("QuickFolders by Voltura AB - " + Application.ProductVersion)
        {
            Image = ThemeHelpers.GetThemeImage("folder"),
            Tag = "folder",
            Enabled = false
        };

        menu.Items.Add(header);
        int added = 1;

        for (int i = 0; i < files.Length; i++)
        {
            FileInfo file = files[i];

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

                FolderMenuItem item = new FolderMenuItem(path, path)
                {
                    Image = ThemeHelpers.GetThemeImage(added.ToString()),
                    Tag = added.ToString()
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
            Image = ThemeHelpers.GetThemeImage("more"),
            Tag = "more"
        };

        TrackSubMenu(menuRoot, openSubMenus);

        DarkToolStripMenuItem web = new DarkToolStripMenuItem("QuickFolders web site")
        {
            Image = ThemeHelpers.GetThemeImage("link"),
            Tag = "link"
        };

        web.Click += OnWebClick;

        menuRoot.DropDownItems.Add(web);

        startWithWindows = new DarkToolStripMenuItem("Start with Windows")
        {
            Image = ThemeHelpers.GetThemeImage("bolt"),
            Checked = ProgramHelpers.StartWithWindows,
            CheckOnClick = true
        };

        DarkToolStripMenuItem folderAction = new DarkToolStripMenuItem("Folder Action...")
        {
            Image = ThemeHelpers.GetThemeImage("folder"),
            Tag = "folder"
        };

        folderAction.Click += OnFolderActionClick;

        menuRoot.DropDownItems.Add(folderAction);

        setDefaultAction = new DarkToolStripMenuItem("Set Default Folder Action")
        {
            Image = ThemeHelpers.GetThemeImage("exit"),
            Tag = "exit"
        };

        setDefaultAction.Click += OnSetDefaultActionClick;

        menuRoot.DropDownItems.Add(setDefaultAction);

        menuRoot.DropDownOpening += OnMenuRootDropDownOpening;

        startWithWindows.CheckedChanged += OnStartWithWindowsCheckedChanged;

        menuRoot.DropDownItems.Add(startWithWindows);

        DarkToolStripMenuItem theme = new DarkToolStripMenuItem("Theme")
        {
            Image = ThemeHelpers.GetThemeImage("theme"),
            Tag = "theme"
        };

        TrackSubMenu(theme, openSubMenus);

        systemTheme = new DarkToolStripMenuItem(Theme.System.ToString())
        {
            Image = ThemeHelpers.GetThemeImage("system"),
            Tag = "system",
            Checked = ThemeHelpers.Config.AppTheme == Theme.System
        };

        darkTheme = new DarkToolStripMenuItem(Theme.Dark.ToString())
        {
            Image = ThemeHelpers.GetThemeImage("darkmode"),
            Tag = "darkmode",
            Checked = ThemeHelpers.Config.AppTheme == Theme.Dark
        };

        lightTheme = new DarkToolStripMenuItem(Theme.Light.ToString())
        {
            Image = ThemeHelpers.GetThemeImage("lightmode"),
            Tag = "lightmode",
            Checked = ThemeHelpers.Config.AppTheme == Theme.Light
        };

        darkTheme.Click += OnDarkThemeClick;
        lightTheme.Click += OnLightThemeClick;
        systemTheme.Click += OnSystemThemeClick;

        theme.DropDownItems.Add(systemTheme);
        theme.DropDownItems.Add(darkTheme);
        theme.DropDownItems.Add(lightTheme);

        menuRoot.DropDownItems.Add(theme);

        DarkToolStripMenuItem fontSize = new DarkToolStripMenuItem("Font Size")
        {
            Image = ThemeHelpers.GetThemeImage("system"),
            Tag = "system"
        };

        TrackSubMenu(fontSize, openSubMenus);

        smallFont = new DarkToolStripMenuItem(FontSize.Small.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Small
        };

        mediumFont = new DarkToolStripMenuItem(FontSize.Medium.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Medium
        };

        largeFont = new DarkToolStripMenuItem(FontSize.Large.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Large
        };

        smallFont.Click += OnSmallFontClick;
        mediumFont.Click += OnMediumFontClick;
        largeFont.Click += OnLargeFontClick;

        fontSize.DropDownItems.Add(smallFont);
        fontSize.DropDownItems.Add(mediumFont);
        fontSize.DropDownItems.Add(largeFont);

        menuRoot.DropDownItems.Add(fontSize);

        DarkToolStripMenuItem exit = new DarkToolStripMenuItem("Exit")
        {
            Image = ThemeHelpers.GetThemeImage("x"),
            Tag = "x"
        };

        exit.Click += OnExitClick;

        menuRoot.DropDownItems.Add(exit);

        menu.Items.Add(menuRoot);

        menu.Font = ThemeHelpers.GetScaledMenuFont();
        ThemeHelpers.ApplyTheme(menu);

        menu.Show(TaskbarMenuPositioner.GetMenuPosition(menu.Size));
        menu.Focus();

        isShowingMenu = false;
        mouseCheckTimer.Start();
    }

    private static void OnMouseCheckTimerTick(object sender, EventArgs e)
    {
        Point cursor = Cursor.Position;

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
    }

    private static void OnAutoCloseTimerTick(object sender, EventArgs e)
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
    }

    private static void OnMenuClosed(object sender, ToolStripDropDownClosedEventArgs e)
    {
        mouseCheckTimer.Stop();
        autoCloseTimer.Stop();

        foreach (var submenu in openSubMenus)
        {
            submenu.Opened -= OnSubMenuOpenedInternal;
            submenu.Closed -= OnSubMenuClosedInternal;
        }

        openSubMenus.Clear();
    }


    private static void OnWebClick(object sender, EventArgs e)
    {
        try
        {
            Process.Start("https://voltura.github.io/QuickFolders");
        }
        catch
        {
        }
    }

    private static void OnFolderActionClick(object sender, EventArgs e)
    {
        string current = ThemeHelpers.Config.FolderActionCommand ?? "";
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
            }

            if (input == current)
            {
                return;
            }

            ThemeHelpers.Config.FolderActionCommand = input;
            ThemeHelpers.Config.Save();

            return;
        }
    }

    private static void OnSetDefaultActionClick(object sender, EventArgs e)
    {
        if (MessageBox.Show("Reset the folder action to default (no custom command)?", "Confirm Reset", MessageBoxButtons.YesNo) == DialogResult.Yes)
        {
            ThemeHelpers.Config.FolderActionCommand = null;
            ThemeHelpers.Config.Save();
        }
    }

    private static void OnMenuRootDropDownOpening(object sender, EventArgs e)
    {
        setDefaultAction.Enabled = !string.IsNullOrEmpty(ThemeHelpers.Config.FolderActionCommand);
    }

    private static void OnStartWithWindowsCheckedChanged(object sender, EventArgs e)
    {
        ProgramHelpers.StartWithWindows = startWithWindows.Checked;
    }

    private static void OnDarkThemeClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.AppTheme = Theme.Dark;
        ThemeHelpers.Config.Save();
        ThemeHelpers.ApplyTheme(menu);
        systemTheme.Checked = false;
        lightTheme.Checked = false;
        darkTheme.Checked = true;
    }

    private static void OnLightThemeClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.AppTheme = Theme.Light;
        ThemeHelpers.Config.Save();
        ThemeHelpers.ApplyTheme(menu);
        systemTheme.Checked = false;
        lightTheme.Checked = true;
        darkTheme.Checked = false;
    }

    private static void OnSystemThemeClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.AppTheme = Theme.System;
        ThemeHelpers.Config.Save();
        ThemeHelpers.ApplyTheme(menu);
        systemTheme.Checked = true;
        lightTheme.Checked = false;
        darkTheme.Checked = false;
    }

    private static void OnSmallFontClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.MenuFontSize = FontSize.Small;
        ThemeHelpers.Config.Save();
        smallFont.Checked = true;
        mediumFont.Checked = false;
        largeFont.Checked = false;
    }

    private static void OnMediumFontClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.MenuFontSize = FontSize.Medium;
        ThemeHelpers.Config.Save();
        smallFont.Checked = false;
        mediumFont.Checked = true;
        largeFont.Checked = false;
    }

    private static void OnLargeFontClick(object sender, EventArgs e)
    {
        ThemeHelpers.Config.MenuFontSize = FontSize.Large;
        ThemeHelpers.Config.Save();
        smallFont.Checked = false;
        mediumFont.Checked = false;
        largeFont.Checked = true;
    }

    private static void OnExitClick(object sender, EventArgs e)
    {
        icon.Visible = false;
        Application.Exit();
    }

    private static void OpenFolder(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(ThemeHelpers.Config.FolderActionCommand))
            {
                string command = ThemeHelpers.Config.FolderActionCommand.Replace("%1", "\"" + path + "\"");
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
    }

    private static void TrackSubMenu(DarkToolStripMenuItem item, List<ToolStripDropDown> submenus)
    {
        item.DropDownOpened += (s, e) => OnSubMenuOpened(item, submenus);
        item.DropDownClosed += (s, e) => OnSubMenuClosed(item, submenus);
    }

    private static void OnSubMenuOpenedInternal(object sender, EventArgs e)
    {
        var dropdown = sender as ToolStripDropDown;
        if (dropdown != null && !openSubMenus.Contains(dropdown))
        {
            openSubMenus.Add(dropdown);
        }
    }

    private static void OnSubMenuClosedInternal(object sender, EventArgs e)
    {
        var dropdown = sender as ToolStripDropDown;
        if (dropdown != null)
        {
            openSubMenus.Remove(dropdown);
        }
    }

    private static void OnSubMenuOpened(DarkToolStripMenuItem item, List<ToolStripDropDown> submenus)
    {
        if (item.DropDown != null && !submenus.Contains(item.DropDown))
        {
            submenus.Add(item.DropDown);
        }
    }

    private static void OnSubMenuClosed(DarkToolStripMenuItem item, List<ToolStripDropDown> submenus)
    {
        if (item.DropDown != null)
        {
            submenus.Remove(item.DropDown);
        }
    }



    private static void DisposeMenuItems(ToolStripItemCollection items)
    {
        for (int i = 0; i < items.Count; i++)
        {
            ToolStripItem item = items[i];

            ToolStripMenuItem menuItem = item as ToolStripMenuItem;
            if (menuItem != null)
            {
                DisposeMenuItems(menuItem.DropDownItems);
            }

            item.Dispose();
        }

        items.Clear();
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
