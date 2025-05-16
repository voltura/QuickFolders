using System;
using System.Drawing;
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

    public static Point GetMenuPosition(Size menuSize)
    {
        Point position = Cursor.Position;

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
