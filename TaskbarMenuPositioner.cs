using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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

    public static Point GetMenuPosition(Size menuSize, NotifyIcon icon, Icon folderIcon, Icon pinkIcon)
    {
        Point position = Cursor.Position;
        icon.Icon = pinkIcon;
        WaitWithDoEvents(100);
        Rectangle? trayIconAreaBounds = GetTrayIconAreaBounds();
        using (Bitmap bmp = CaptureTrayIconArea(trayIconAreaBounds))
        {
            icon.Icon = folderIcon;
            Point? center = FindBoundingBoxCenter(bmp);

            if (center.HasValue)
            {
                Rectangle trayBounds = GetTrayIconAreaBounds().Value;
                Point screenPos = new Point(trayBounds.Left + center.Value.X, trayBounds.Top + center.Value.Y);
                position.X = screenPos.X;
            }
        }
        if (position.X < trayIconAreaBounds.Value.Left)
        {
            position.X = trayIconAreaBounds.Value.Left + 16;
        }
        if (position.X > trayIconAreaBounds.Value.Right)
        {
            position.X = trayIconAreaBounds.Value.Right - menuSize.Width;
        }

        TaskbarPosition taskbarPos;
        Rectangle taskbarRect;
        if (GetTaskbarPosition(out taskbarPos, out taskbarRect))
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

    private static void WaitWithDoEvents(int milliseconds)
    {
        var end = Environment.TickCount + milliseconds;
        while (Environment.TickCount < end)
        {
            Application.DoEvents();
        }
    }


    public static Point? FindBoundingBoxCenter(Bitmap bmp)
    {
        Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
        BitmapData data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        int minX = bmp.Width;
        int maxX = -1;
        int minY = bmp.Height;
        int maxY = -1;

        unsafe
        {
            byte* ptr = (byte*)data.Scan0;

            for (int y = 0; y < data.Height; y++)
            {
                byte* row = ptr + (y * data.Stride);
                for (int x = 0; x < data.Width; x++)
                {
                    byte b = row[x * 4 + 0];
                    byte g = row[x * 4 + 1];
                    byte r = row[x * 4 + 2];

                    if (r == 255 && g == 20 && b == 147) // #FF1493
                    {
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
            }
        }

        bmp.UnlockBits(data);

        if (maxX < minX || maxY < minY)
        {
            return null;
        }

        return new Point((minX + maxX) / 2, (minY + maxY) / 2);
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

    public static Bitmap CaptureTrayIconArea(Rectangle? trayBounds)
    {
        if (!trayBounds.HasValue)
        {
            return null;
        }

        Rectangle bounds = trayBounds.Value;
        Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);

        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
        }

        return bitmap;
    }

    public static Rectangle? GetTrayIconAreaBounds()
    {
        IntPtr hShellTrayWnd = FindWindow("Shell_TrayWnd", null);
        if (hShellTrayWnd == IntPtr.Zero)
        {
            return null;
        }

        IntPtr hTrayNotifyWnd = FindWindowEx(hShellTrayWnd, IntPtr.Zero, "TrayNotifyWnd", null);
        if (hTrayNotifyWnd == IntPtr.Zero)
        {
            return null;
        }

        RECT rect;
        if (GetWindowRect(hTrayNotifyWnd, out rect))
        {
            return rect.ToRectangle();
        }

        return null;
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
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
