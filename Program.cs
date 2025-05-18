using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyVersion("1.0.1.2")]
[assembly: AssemblyFileVersion("1.0.1.2")]
[assembly: AssemblyInformationalVersion("v1.0.1.2")]
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

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
    private static readonly object lockObject = new object();
    private static CustomToolStripDropDownMenu menu;
    private static System.Windows.Forms.Timer mouseCheckTimer;
    private static System.Windows.Forms.Timer autoCloseTimer;
    private static CustomToolStripMenuItem setDefaultAction;
    private static CustomToolStripMenuItem startWithWindows;
    private static CustomToolStripMenuItem systemTheme;
    private static CustomToolStripMenuItem lightTheme;
    private static CustomToolStripMenuItem darkTheme;
    private static CustomToolStripMenuItem smallFont;
    private static CustomToolStripMenuItem mediumFont;
    private static CustomToolStripMenuItem largeFont;
    private static NotifyIcon icon;
    private static bool isShowingMenu = false;
    private static bool inAutoClose = false;
    private static readonly FolderMenuItem[] recentFolderItems = new FolderMenuItem[5];
    private static Icon pinkIcon = null;
    private static Icon folderIcon = null;
    private static HotKey hotKey;
    private static System.Windows.Forms.Timer hoverTimer;
    private static Point lastHoverPoint;
    private static DateTime hoverStartTime;
    private static DateTime lastInputTime = DateTime.UtcNow;

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
        folderIcon = Icon.ExtractAssociatedIcon(Environment.ExpandEnvironmentVariables(@"%SystemRoot%\explorer.exe"));
        pinkIcon = CreateMarkedIcon(folderIcon);

        // create icon
        // -----------
        icon = new NotifyIcon
        {
            Icon = folderIcon,
            Visible = true,
            Text = null
        };

        // create menu
        // -----------
        menu = new CustomToolStripDropDownMenu
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

        // create menu items
        // -----------------

        // header menu item
        CustomToolStripMenuItem header = new CustomToolStripMenuItem("QuickFolders by Voltura AB - " + Application.ProductVersion)
        {
            Image = ThemeHelpers.GetThemeImage("folder"),
            Tag = "folder",
            Enabled = false
        };
        menu.Items.Add(header);

        // add 5 placeholder items for folders
        for (int i = 0; i < 5; i++)
        {
            recentFolderItems[i] = new FolderMenuItem("Loading...", "")
            {
                Image = ThemeHelpers.GetThemeImage((i + 1).ToString()),
                Tag = (i + 1).ToString()
            };
            recentFolderItems[i].Click += OnRecentFolderItemClick;
            menu.Items.Add(recentFolderItems[i]);
        }

        // add root menu
        menu.Items.Add(BuildRootMenu());

        WarmUpAllDropDowns(menu.Items);

        // create timers
        mouseCheckTimer = new System.Windows.Forms.Timer
        {
            Interval = 100
        };

        autoCloseTimer = new System.Windows.Forms.Timer
        {
            Interval = 1500
        };

        hoverTimer = new System.Windows.Forms.Timer
        {
            Interval = 100
        };

        mouseCheckTimer.Tick += OnMouseCheckTimerTick;
        autoCloseTimer.Tick += OnAutoCloseTimerTick;
        hoverTimer.Tick += OnHoverTimerTick;
        autoCloseTimer.Start();

        // define actions on mouse events
        menu.MouseEnter += OnMouseEnteredMenu;
        menu.MouseLeave += OnMouseLeftMenu;
        icon.MouseClick += OnIconClick;
        icon.MouseMove += OnIconHoverMove;

        // add shortcuts 1-5
        menu.PreviewKeyDown += OnMenuPreviewKeyDown;
        menu.KeyDown += OnMenuKeyDown;

        // add hotkey CTRL+SHIFT+SPACE
        hotKey = new HotKey(ShowMenu);

        // run application
        Application.Run(new ApplicationContext());
    }

    private static void OnIconClick(object sender, MouseEventArgs e)
    {
        ShowMenu();
    }

    private static void OnIconHoverMove(object sender, MouseEventArgs e)
    {
        lastHoverPoint = Cursor.Position;
        hoverStartTime = DateTime.UtcNow;

        if (!hoverTimer.Enabled)
        {
            hoverTimer.Start();
        }
    }

    private static void OnHoverTimerTick(object sender, EventArgs e)
    {
        if (menu.Visible || isShowingMenu)
        {
            hoverTimer.Stop();
            return;
        }

        TimeSpan hoverDuration = DateTime.UtcNow - hoverStartTime;

        if (hoverDuration.TotalMilliseconds >= 400)
        {
            Point current = Cursor.Position;

            if (Math.Abs(current.X - lastHoverPoint.X) <= 2 &&
                Math.Abs(current.Y - lastHoverPoint.Y) <= 2)
            {
                hoverTimer.Stop();
                ShowMenu();
            }
            else
            {
                lastHoverPoint = current;
                hoverStartTime = DateTime.UtcNow;
            }
        }
    }

    private static void ShowMenu()
    {
        if (isShowingMenu || menu == null || menu.Visible)
        {
            return;
        }

        lock (lockObject)
        {
            isShowingMenu = true;
            UpdateRecentFolderItems();
            menu.Font = ThemeHelpers.GetScaledMenuFont();
            ThemeHelpers.ApplyTheme(menu);
            
            menu.Show(TaskbarMenuPositioner.GetMenuPosition(menu.Size, icon, folderIcon, pinkIcon));
            menu.Focus();
            menu.Items[1].Select();

            Application.DoEvents();

            if (menu.Handle != IntPtr.Zero)
            {
                SetForegroundWindow(menu.Handle);
            }

            isShowingMenu = false;
        }

        lastInputTime = DateTime.UtcNow;
        mouseCheckTimer.Start();
        autoCloseTimer.Start();
    }

    private static void OnMouseLeftMenu(object sender, EventArgs e)
    {
        lastInputTime = DateTime.UtcNow;
    }

    private static void OnMenuPreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
    {
        ResetAutoClose();

        if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D5)
        {
            e.IsInputKey = true;
        }
    }

    private static void ResetAutoClose()
    {
        lastInputTime = DateTime.UtcNow;

        if (autoCloseTimer != null)
        {
            autoCloseTimer.Stop();
            autoCloseTimer.Start();
        }
    }

    private static void OnMenuKeyDown(object sender, KeyEventArgs e)
    {
        int index = -1;

        if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D5)
        {
            index = e.KeyCode - Keys.D1;
        }
        else if (e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad5)
        {
            index = e.KeyCode - Keys.NumPad1;
        }

        if (index >= 0 && index < recentFolderItems.Length)
        {
            var item = recentFolderItems[index];

            if (item != null && item.Enabled && !string.IsNullOrEmpty(item.FolderPath))
            {
                OpenFolder(item.FolderPath);
                menu.Close();
                e.Handled = true;
            }
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    public static Icon CreateMarkedIcon(Icon baseIcon)
    {
        using (Bitmap bmp = baseIcon.ToBitmap())
        using (Graphics g = Graphics.FromImage(bmp))
        {
            int dotSize = 3;
            int centerX = (bmp.Width - dotSize) / 2;
            int centerY = (bmp.Height - dotSize) / 2;

            Rectangle dotArea = new Rectangle(centerX, centerY, dotSize, dotSize);

            using (Brush brush = new SolidBrush(Color.FromArgb(255, 20, 147)))
            {
                g.FillRectangle(brush, dotArea);
            }

            IntPtr hIcon = bmp.GetHicon();

            using (var icon = Icon.FromHandle(hIcon))
            {
                Icon clone = (Icon)icon.Clone();
                DestroyIcon(hIcon);
                return clone;
            }
        }
    }

    private static void OnRecentFolderItemClick(object sender, EventArgs e)
    {
        var item = sender as FolderMenuItem;

        if (item != null && !string.IsNullOrEmpty(item.FolderPath))
        {
            OpenFolder(item.FolderPath);
        }
    }

    
    private static void OnMouseEnteredMenu(object sender, EventArgs e)
    {
        lastInputTime = DateTime.UtcNow;
    }

    private static void UpdateRecentFolderItems()
    {
        List<string> recentPaths = RecentFolderProvider.GetRecentFolders(5);

        int updated = 0;

        for (int i = 0; i < recentPaths.Count && updated < 5; i++)
        {
            string path = recentPaths[i];

            recentFolderItems[updated].Text = path;
            recentFolderItems[updated].FolderPath = path;
            recentFolderItems[updated].Enabled = true;

            updated++;
        }

        for (int i = updated; i < 5; i++)
        {
            recentFolderItems[i].Text = "";
            recentFolderItems[i].FolderPath = "";
            recentFolderItems[i].Enabled = false;
        }
    }

    private static CustomToolStripMenuItem BuildRootMenu()
    {
        CustomToolStripMenuItem menuRoot = new CustomToolStripMenuItem("Menu")
        {
            Image = ThemeHelpers.GetThemeImage("more"),
            Tag = "more"
        };

        CustomToolStripMenuItem web = new CustomToolStripMenuItem("QuickFolders web site")
        {
            Image = ThemeHelpers.GetThemeImage("link"),
            Tag = "link"
        };

        web.Click += OnWebClick;

        menuRoot.DropDownItems.Add(web);

        bool doStartWithWindows = ProgramHelpers.StartWithWindows;

        startWithWindows = new CustomToolStripMenuItem("Start with Windows")
        {
            Image = doStartWithWindows ? null : ThemeHelpers.GetThemeImage("bolt"),
            Checked = doStartWithWindows,
            CheckOnClick = true,
            Tag = "bolt"
        };

        CustomToolStripMenuItem folderAction = new CustomToolStripMenuItem("Folder Action...")
        {
            Image = ThemeHelpers.GetThemeImage("folder"),
            Tag = "folder"
        };

        folderAction.Click += OnFolderActionClick;

        menuRoot.DropDownItems.Add(folderAction);

        setDefaultAction = new CustomToolStripMenuItem("Set Default Folder Action")
        {
            Image = ThemeHelpers.GetThemeImage("exit"),
            Tag = "exit"
        };

        setDefaultAction.Click += OnSetDefaultActionClick;

        menuRoot.DropDownItems.Add(setDefaultAction);

        menuRoot.DropDownOpening += OnMenuRootDropDownOpening;

        startWithWindows.CheckedChanged += OnStartWithWindowsCheckedChanged;

        menuRoot.DropDownItems.Add(startWithWindows);

        CustomToolStripMenuItem theme = new CustomToolStripMenuItem("Theme")
        {
            Image = ThemeHelpers.GetThemeImage("theme"),
            Tag = "theme"
        };

        systemTheme = new CustomToolStripMenuItem(Theme.System.ToString())
        {
            Image = ThemeHelpers.GetThemeImage("system"),
            Tag = "system",
            Checked = ThemeHelpers.Config.AppTheme == Theme.System
        };

        darkTheme = new CustomToolStripMenuItem(Theme.Dark.ToString())
        {
            Image = ThemeHelpers.GetThemeImage("darkmode"),
            Tag = "darkmode",
            Checked = ThemeHelpers.Config.AppTheme == Theme.Dark
        };

        lightTheme = new CustomToolStripMenuItem(Theme.Light.ToString())
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
        theme.DropDownOpening += OnMenuOpening;
        theme.MouseEnter += OnMenuItemWithChildrenMouseEnter;
        menuRoot.DropDownItems.Add(theme);

        CustomToolStripMenuItem fontSize = new CustomToolStripMenuItem("Font Size")
        {
            Image = ThemeHelpers.GetThemeImage("system"),
            Tag = "system"
        };

        smallFont = new CustomToolStripMenuItem(FontSize.Small.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Small
        };

        mediumFont = new CustomToolStripMenuItem(FontSize.Medium.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Medium
        };

        largeFont = new CustomToolStripMenuItem(FontSize.Large.ToString())
        {
            Checked = ThemeHelpers.Config.MenuFontSize == FontSize.Large
        };

        smallFont.Click += OnSmallFontClick;
        mediumFont.Click += OnMediumFontClick;
        largeFont.Click += OnLargeFontClick;

        fontSize.DropDownItems.Add(smallFont);
        fontSize.DropDownItems.Add(mediumFont);
        fontSize.DropDownItems.Add(largeFont);
        fontSize.MouseEnter += OnMenuItemWithChildrenMouseEnter;
        fontSize.DropDownOpening += OnMenuOpening;
        menuRoot.DropDownItems.Add(fontSize);

        CustomToolStripMenuItem exit = new CustomToolStripMenuItem("Exit")
        {
            Image = ThemeHelpers.GetThemeImage("x"),
            Tag = "x"
        };

        exit.Click += OnExitClick;

        menuRoot.DropDownItems.Add(exit);

        menuRoot.DropDownOpening += OnMenuOpening;

        return menuRoot;
    }

    private static void OnMenuOpening(object sender, EventArgs e)
    {
        ResetAutoClose();
    }

    private static void OnMenuItemWithChildrenMouseEnter(object sender, EventArgs e)
    {
        ToolStripMenuItem current = sender as ToolStripMenuItem;

        if (current == null)
        {
            return;
        }

        if (!current.DropDown.Visible && current.HasDropDownItems)
        {
            current.ShowDropDown();
        }

        CloseUnrelatedSubmenus(current);
    }

    private static void CloseUnrelatedSubmenus(ToolStripMenuItem keepOpenRoot)
    {
        if (menu == null)
        {
            return;
        }

        foreach (ToolStripItem item in menu.Items)
        {
            ToolStripMenuItem topItem = item as ToolStripMenuItem;

            if (topItem == null)
            {
                continue;
            }

            if (!topItem.HasDropDownItems)
            {
                continue;
            }

            if (topItem == keepOpenRoot || IsDescendant(topItem, keepOpenRoot))
            {
                continue;
            }

            if (topItem.DropDown.Visible)
            {
                topItem.HideDropDown();
            }
        }
    }

    private static bool IsDescendant(ToolStripMenuItem root, ToolStripMenuItem target)
    {
        if (root == null || target == null)
        {
            return false;
        }

        foreach (ToolStripItem item in root.DropDownItems)
        {
            ToolStripMenuItem child = item as ToolStripMenuItem;

            if (child == null)
            {
                continue;
            }

            if (child == target)
            {
                return true;
            }

            if (IsDescendant(child, target))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsCursorOverAnyVisibleSubmenu(Control parent, Point cursorPos)
    {
        ToolStripDropDown dropDown = parent as ToolStripDropDown;

        if (dropDown != null && dropDown.Visible && dropDown.Bounds.Contains(cursorPos))
        {
            return true;
        }

        ToolStripDropDownMenu menu = parent as ToolStripDropDownMenu;

        if (menu == null)
        {
            return false;
        }

        foreach (ToolStripItem item in menu.Items)
        {
            ToolStripMenuItem menuItem = item as ToolStripMenuItem;

            if (menuItem != null && menuItem.DropDown != null && menuItem.DropDown.Visible)
            {
                if (IsCursorOverAnyVisibleSubmenu(menuItem.DropDown, cursorPos))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static void OnMouseCheckTimerTick(object sender, EventArgs e)
    {
        if (!menu.Visible || inAutoClose == true)
        {
            return;
        }

        Point cursorPos = Cursor.Position;

        if (menu.Bounds.Contains(cursorPos) || IsCursorOverAnyVisibleSubmenu(menu, cursorPos))
        {
            lastInputTime = DateTime.UtcNow;
            return;
        }
    }

    private static void OnAutoCloseTimerTick(object sender, EventArgs e)
    {
        if (!menu.Visible)
        {
            return;
        }

        Point cursorPos = Cursor.Position;

        if (menu.Bounds.Contains(cursorPos) || IsCursorOverAnyVisibleSubmenu(menu, cursorPos))
        {
            lastInputTime = DateTime.UtcNow;

            return;
        }

        TimeSpan idle = DateTime.UtcNow - lastInputTime;

        if (idle.TotalSeconds >= 3.0)
        {
            inAutoClose = true;

            try
            {
                menu.Close();
            }
            finally
            {
                inAutoClose = false;
            }
        }
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
        OnMenuOpening(sender, e);
    }

    private static void OnStartWithWindowsCheckedChanged(object sender, EventArgs e)
    {
        ProgramHelpers.StartWithWindows = startWithWindows.Checked;
    }

    private static void SetTheme(Theme theme)
    {
        ThemeHelpers.Config.AppTheme = theme;
        ThemeHelpers.Config.Save();
        ThemeHelpers.ApplyTheme(menu);
        systemTheme.Checked = (theme == Theme.System);
        lightTheme.Checked = (theme == Theme.Light);
        darkTheme.Checked = (theme == Theme.Dark);
    }

    private static void OnDarkThemeClick(object sender, EventArgs e)
    {
        SetTheme(Theme.Dark);
    }

    private static void OnLightThemeClick(object sender, EventArgs e)
    {
        SetTheme(Theme.Light);
    }

    private static void OnSystemThemeClick(object sender, EventArgs e)
    {
        SetTheme(Theme.System);
    }

    private static void SetFontSize(FontSize fontSize)
    {
        ThemeHelpers.Config.MenuFontSize = fontSize;
        ThemeHelpers.Config.Save();
        smallFont.Checked = (fontSize == FontSize.Small);
        mediumFont.Checked = (fontSize == FontSize.Medium);
        largeFont.Checked = (fontSize == FontSize.Large);
    }

    private static void OnSmallFontClick(object sender, EventArgs e)
    {
        SetFontSize(FontSize.Small);
    }

    private static void OnMediumFontClick(object sender, EventArgs e)
    {
        SetFontSize(FontSize.Medium);
    }

    private static void OnLargeFontClick(object sender, EventArgs e)
    {
        SetFontSize(FontSize.Large);
    }

    private static void OnExitClick(object sender, EventArgs e)
    {
        if (mouseCheckTimer != null)
        {
            mouseCheckTimer.Stop();
            mouseCheckTimer.Tick -= OnMouseCheckTimerTick;
            mouseCheckTimer.Dispose();
            mouseCheckTimer = null;
        }

        if (autoCloseTimer != null)
        {
            autoCloseTimer.Stop();
            autoCloseTimer.Tick -= OnAutoCloseTimerTick;
            autoCloseTimer.Dispose();
            autoCloseTimer = null;
        }

        if (icon != null)
        {
            icon.Visible = false;
            icon.Dispose();
            icon = null;
        }

        if (folderIcon != null)
        {
            folderIcon.Dispose();
            folderIcon = null;
        }

        if (pinkIcon != null)
        {
            pinkIcon.Dispose();
            pinkIcon = null;
        }

        if (hotKey != null)
        {
            hotKey.Dispose();
            hotKey = null;
        }

        ThemeHelpers.DisposeCachedImages();
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

    private static void WarmUpAllDropDowns(ToolStripItemCollection items)
    {
        for (int i = 0; i < items.Count; i++)
        {
            ToolStripMenuItem menuItem = items[i] as ToolStripMenuItem;

            if (menuItem != null && menuItem.HasDropDownItems)
            {
                Control dummy = menuItem.DropDown;
                
                if (!menuItem.DropDown.Visible)
                {
                    menuItem.ShowDropDown();
                    menuItem.HideDropDown();
                }

                WarmUpAllDropDowns(menuItem.DropDownItems);
            }
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
}
