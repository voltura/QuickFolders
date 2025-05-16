using Microsoft.Win32;
using System.Windows.Forms;

internal static class ProgramHelpers
{

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
}