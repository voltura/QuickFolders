using Microsoft.Win32;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

public class DarkToolStripDropDownMenu : ToolStripDropDownMenu
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

public class DarkToolStripMenuItem : ToolStripMenuItem
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

public class DarkColorTable : ProfessionalColorTable
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

public static class ThemeHelpers
{
    public class DarkMenuRenderer : ToolStripProfessionalRenderer
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

    public static readonly ToolStripRenderer DarkRenderer = new DarkMenuRenderer();
    public static readonly ToolStripRenderer DefaultRenderer = new ToolStripProfessionalRenderer();
    public static Config Config;

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

    public static void ApplyTheme(DarkToolStripDropDownMenu menu)
    {
        string suffix = Config.AppTheme == Theme.Dark ? "_dark" : "";
        Theme realSystemTheme = Config.AppTheme == Theme.Dark ? Theme.Dark : Theme.Light;

        if (Config.AppTheme == Theme.System)
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

    public static void ApplyThemeToMenuItem(ToolStripItem item, string suffix, Theme realSystemTheme)
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

    public static Font GetScaledMenuFont()
    {
        float scaleFactor;

        switch (Config.MenuFontSize)
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

    public static Image GetThemeImage(string baseName, string overrideSuffix = null)
    {
        string suffix = Config.AppTheme == Theme.Dark ? "_dark" : "";
        string resourceName = "QuickFolders.Resources." + baseName + (overrideSuffix ?? suffix) + ".png";

        using (Stream stream = typeof(Program).Assembly.GetManifestResourceStream(resourceName))
        {
            return stream == null ? null : Image.FromStream(stream);
        }
    }
}