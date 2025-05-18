using Microsoft.Win32;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

public class CustomToolStripDropDownMenu : ToolStripDropDownMenu
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

public class CustomToolStripMenuItem : ToolStripMenuItem
{
    public CustomToolStripMenuItem() : base() { }

    public CustomToolStripMenuItem(string text) : base(text) { }

    protected override ToolStripDropDown CreateDefaultDropDown()
    {
        CustomToolStripDropDownMenu dropDown = new CustomToolStripDropDownMenu
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

public class LightColorTable : ProfessionalColorTable
{
    public override Color MenuItemSelected
    {
        get { return Color.FromArgb(230, 230, 230); } // fallback solid
    }

    public override Color MenuItemSelectedGradientBegin
    {
        get { return Color.FromArgb(250, 250, 250); }
    }

    public override Color MenuItemSelectedGradientEnd
    {
        get { return Color.FromArgb(200, 200, 200); }
    }

    public override Color MenuItemBorder
    {
        get { return Color.LightGray; }
    }

    public override Color ToolStripBorder
    {
        get { return Color.LightGray; }
    }

    public override Color MenuBorder
    {
        get { return Color.LightGray; }
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
            rect.Y -= 1;
            rect.Height += 2;

            if (e.Item.Selected)
            {
                using (var brush = new LinearGradientBrush(
                    rect,
                    Color.FromArgb(70, 70, 70),
                    Color.FromArgb(50, 50, 50),
                    LinearGradientMode.Vertical))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }
            }
            else if (!e.Item.Enabled)
            {
                using (Brush b = new SolidBrush(Color.FromArgb(32, 32, 32)))
                {
                    e.Graphics.FillRectangle(b, rect);
                }
            }
            else
            {
                using (Brush b = new SolidBrush(Color.FromArgb(32, 32, 32)))
                {
                    e.Graphics.FillRectangle(b, rect);
                }
            }
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            int size = GetScaledArrowSize(e.Graphics.DpiX);
            Rectangle rect = e.ArrowRectangle;
            Point middle = new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2);

            Point[] arrow =
            {
                new Point(middle.X - size / 2, middle.Y - size / 2),
                new Point(middle.X + size / 2, middle.Y),
                new Point(middle.X - size / 2, middle.Y + size / 2)
            };

            Color arrowColor = (!e.Item.Enabled) ? Color.DimGray :
                               (e.Item.Selected) ? Color.White : Color.Gainsboro;

            using (Brush brush = new SolidBrush(arrowColor))
            using (Pen edge = new Pen(Color.FromArgb(60, 60, 60)))
            {
                e.Graphics.FillPolygon(brush, arrow);
                e.Graphics.DrawPolygon(edge, arrow);
            }
        }

        protected override void OnRenderImageMargin(ToolStripRenderEventArgs e)
        {
        }

        protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
        {
            float dpi = e.Graphics.DpiX;
            Rectangle rect = GetCenteredSquare(e.ImageRectangle, dpi, 16);
            Rectangle insetRect = new Rectangle(rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

            Color backColor = Color.FromArgb(80, 80, 80);       // background fill
            Color borderColor = Color.FromArgb(140, 140, 140);  // border
            Color tickColor = Color.White;                      // checkmark

            using (Brush b = new SolidBrush(backColor))
            using (Pen border = new Pen(borderColor))
            using (Pen tick = new Pen(tickColor, 2f * dpi / 96f))
            using (GraphicsPath path = CreateRoundedRectangle(insetRect, 3)) // corner radius = 3
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillPath(b, path);
                e.Graphics.DrawPath(border, path);

                DrawCheckMark(e.Graphics, insetRect, tick);
            }
        }

        private void DrawCheckMark(Graphics g, Rectangle bounds, Pen pen)
        {
            int inset = bounds.Width / 4;
            Point p1 = new Point(bounds.Left + inset, bounds.Top + bounds.Height / 2);
            Point p2 = new Point(bounds.Left + bounds.Width / 2 - 1, bounds.Bottom - inset - 1);
            Point p3 = new Point(bounds.Right - inset - 1, bounds.Top + inset);

            g.DrawLines(pen, new Point[] { p1, p2, p3 });
        }

        private Rectangle GetCenteredSquare(Rectangle target, float dpi, int baseSize)
        {
            int size = (int)(baseSize * dpi / 96f);
            int x = target.Left + (target.Width - size) / 2;
            int y = target.Top + (target.Height - size) / 2;

            return new Rectangle(x + 1, y, size, size);
        }

        private GraphicsPath CreateRoundedRectangle(Rectangle bounds, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;

            Rectangle arc = new Rectangle(bounds.X, bounds.Y, diameter, diameter);
            path.AddArc(arc, 180, 90); // top-left

            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90); // top-right

            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90); // bottom-right

            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90); // bottom-left

            path.CloseFigure();

            return path;
        }

        private int GetScaledArrowSize(float dpi)
        {
            return (int)(8 * dpi / 96f);
        }
    }

    public class LightMenuRenderer : ToolStripProfessionalRenderer
    {
        public LightMenuRenderer() : base(new LightColorTable()) { }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using (Pen p = new Pen(Color.LightGray, 1))
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
            Color textColor = e.Item.Enabled ? Color.Black : Color.Gray;
            TextRenderer.DrawText(e.Graphics, e.Text, e.TextFont, e.TextRectangle, textColor, flags);
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rect = new Rectangle(Point.Empty, e.Item.Size);
            rect.Y -= 1;
            rect.Height += 2;

            if (e.Item.Selected)
            {
                using (var brush = new LinearGradientBrush(
                    rect,
                    Color.FromArgb(255, 255, 255),
                    Color.FromArgb(220, 220, 220),
                    LinearGradientMode.Vertical))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }
            }
            else
            {
                using (Brush brush = new SolidBrush(Color.WhiteSmoke))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }
            }
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            int size = GetScaledArrowSize(e.Graphics.DpiX);
            Rectangle rect = e.ArrowRectangle;
            Point middle = new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2);

            Point[] arrow =
            {
                new Point(middle.X - size / 2, middle.Y - size / 2),
                new Point(middle.X + size / 2, middle.Y),
                new Point(middle.X - size / 2, middle.Y + size / 2)
            };

            Color arrowColor = (!e.Item.Enabled) ? Color.LightGray :
                               (e.Item.Selected) ? Color.Black : Color.DarkGray;

            using (Brush brush = new SolidBrush(arrowColor))
            using (Pen edge = new Pen(Color.FromArgb(220, 220, 220)))
            {
                e.Graphics.FillPolygon(brush, arrow);
                e.Graphics.DrawPolygon(edge, arrow);
            }
        }

        protected override void OnRenderImageMargin(ToolStripRenderEventArgs e)
        {
        }

        protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
        {
            float dpi = e.Graphics.DpiX;
            Rectangle rect = GetCenteredSquare(e.ImageRectangle, dpi, 16);
            Rectangle insetRect = new Rectangle(rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

            Color backColor = Color.White;
            Color borderColor = Color.LightGray;
            Color tickColor = Color.Black;

            using (Brush b = new SolidBrush(backColor))
            using (Pen border = new Pen(borderColor))
            using (Pen tick = new Pen(tickColor, 2f * dpi / 96f))
            using (GraphicsPath path = CreateRoundedRectangle(insetRect, 3))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillPath(b, path);
                e.Graphics.DrawPath(border, path);
                DrawCheckMark(e.Graphics, insetRect, tick);
            }
        }

        private void DrawCheckMark(Graphics g, Rectangle bounds, Pen pen)
        {
            int inset = bounds.Width / 4;
            Point p1 = new Point(bounds.Left + inset, bounds.Top + bounds.Height / 2);
            Point p2 = new Point(bounds.Left + bounds.Width / 2 - 1, bounds.Bottom - inset - 1);
            Point p3 = new Point(bounds.Right - inset - 1, bounds.Top + inset);
            g.DrawLines(pen, new Point[] { p1, p2, p3 });
        }

        private Rectangle GetCenteredSquare(Rectangle target, float dpi, int baseSize)
        {
            int size = (int)(baseSize * dpi / 96f);
            int x = target.Left + (target.Width - size) / 2;
            int y = target.Top + (target.Height - size) / 2;

            return new Rectangle(x + 1, y, size, size);
        }

        private GraphicsPath CreateRoundedRectangle(Rectangle bounds, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;
            Rectangle arc = new Rectangle(bounds.X, bounds.Y, diameter, diameter);
            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();

            return path;
        }

        private int GetScaledArrowSize(float dpi)
        {
            return (int)(8 * dpi / 96f);
        }
    }

    public static readonly ToolStripRenderer DarkRenderer = new DarkMenuRenderer();
    public static readonly ToolStripRenderer LightRenderer = new LightMenuRenderer();

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

    public static void ApplyTheme(CustomToolStripDropDownMenu menu)
    {
        string suffix = Config.AppTheme == Theme.Dark ? "_dark" : "";
        Theme realSystemTheme = Config.AppTheme == Theme.Dark ? Theme.Dark : Theme.Light;

        if (Config.AppTheme == Theme.System)
        {
            realSystemTheme = GetSystemTheme();
        }

        menu.Font = GetScaledMenuFont();
        menu.Renderer = realSystemTheme == Theme.Dark ? DarkRenderer : LightRenderer;

        foreach (ToolStripItem item in menu.Items)
        {
            ApplyThemeToMenuItem(item, suffix, realSystemTheme);
        }
    }

    public static void ApplyThemeToMenuItem(ToolStripItem item, string suffix, Theme realSystemTheme)
    {
        CustomToolStripMenuItem menuItem = item as CustomToolStripMenuItem;

        if (menuItem != null)
        {
            if (menuItem.Tag != null)
            {
                string baseName = menuItem.Tag.ToString();
                menuItem.Image = (menuItem.Checked) ? null : GetThemeImage(baseName, suffix);
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

            menuItem.ImageScaling = ToolStripItemImageScaling.SizeToFit;
            menuItem.TextImageRelation = TextImageRelation.ImageBeforeText;
            menuItem.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;

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

    private static readonly Dictionary<string, Image> _imageCache = new Dictionary<string, Image>();

    public static Image GetThemeImage(string baseName, string overrideSuffix = null)
    {
        string suffix = Config.AppTheme == Theme.Dark ? "_dark" : "";
        string key = baseName + (overrideSuffix ?? suffix);

        lock (_imageCache)
        {
            Image cached;

            if (_imageCache.TryGetValue(key, out cached))
            {
                return cached;
            }

            string resourceName = "QuickFolders.Resources." + key + ".png";

            Stream stream = typeof(Program).Assembly.GetManifestResourceStream(resourceName);

            if (stream == null)
            {
                return null;
            }

            Image img = Image.FromStream(stream);
            _imageCache[key] = img;

            return img;
        }
    }

    public static void DisposeCachedImages()
    {
        lock (_imageCache)
        {
            foreach (var img in _imageCache.Values)
            {
                img.Dispose();
            }

            _imageCache.Clear();
        }
    }
}
