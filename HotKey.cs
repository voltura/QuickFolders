using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotKey : NativeWindow, IDisposable
{
    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 1;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private readonly Action onTriggered;

    public HotKey(Action onTriggered)
    {
        this.onTriggered = onTriggered;
        CreateHandle(new CreateParams());

        bool success = RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.Space);
        if (!success)
        {
#if DEBUG
            MessageBox.Show("Failed to register global hotkey Ctrl+Shift+Space.", "QuickFolders", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
        }
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
        {
            if (onTriggered != null)
            {
                onTriggered();
            }
        }

        base.WndProc(ref m);
    }

    public void Dispose()
    {
        UnregisterHotKey(Handle, HOTKEY_ID);
        DestroyHandle();
    }
}
