sealed class FolderMenuItem : DarkToolStripMenuItem
{
    private readonly string _folderPath;
    public string FolderPath { get { return _folderPath; } }

    public FolderMenuItem(string text, string folderPath)
        : base(text)
    {
        _folderPath = folderPath;
    }
}
