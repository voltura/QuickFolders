sealed class FolderMenuItem : CustomToolStripMenuItem
{
    private string _folderPath;
    public string FolderPath { get { return _folderPath; } set { _folderPath = value; } }

    public FolderMenuItem(string text, string folderPath) : base(text)
    {
        _folderPath = folderPath;
    }
}
