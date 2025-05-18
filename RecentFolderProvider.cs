using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

static class RecentFolderProvider
{
    public static List<string> GetRecentFolders(int maxCount)
    {
        var result = new List<string>();
        FileInfo[] files;

        try
        {
            files = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Recent)).GetFiles("*.lnk");
        }
        catch
        {
            return result;
        }

        Array.Sort(files, (a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));

        for (int i = 0; i < files.Length && result.Count < maxCount; i++)
        {
            IShellLink link = null;
            IPersistFile persist = null;

            try
            {
                link = (IShellLink)new ShellLink();
                persist = (IPersistFile)link;
                persist.Load(files[i].FullName, 0);

                StringBuilder target = new StringBuilder(260);
                link.GetPath(target, target.Capacity, IntPtr.Zero, 0);
                string path = target.ToString();

                if (!string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path) && Directory.Exists(path))
                {
                    result.Add(path);
                }
            }
            catch
            {
            }
            finally
            {
                if (persist != null)
                {
                    Marshal.ReleaseComObject(persist);
                }

                if (link != null)
                {
                    Marshal.ReleaseComObject(link);
                }
            }
        }

        return result;
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
