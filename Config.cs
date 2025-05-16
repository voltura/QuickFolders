using System;
using System.Collections.Generic;
using System.IO;

public class Config
{
    public string FolderActionCommand = string.Empty;

    public Theme AppTheme = Theme.System;

    public FontSize MenuFontSize = FontSize.Medium;

    public static string ConfigFilePath
    {
        get
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "QuickFolders");
            Directory.CreateDirectory(folder);

            return Path.Combine(folder, "QuickFolders.config");
        }
    }

    public static Config Load()
    {
        if (!File.Exists(ConfigFilePath))
        {
            Config newConfig = new Config();
            newConfig.Save();

            return newConfig;
        }

        Config config = new Config();

        string[] lines = File.ReadAllLines(ConfigFilePath);

        foreach (string line in lines)
        {
            string trimmed = line.Trim();

            if (string.IsNullOrEmpty(trimmed))
            {
                continue;
            }

            if (trimmed.StartsWith("#"))
            {
                continue;
            }

            string[] parts = trimmed.Split(new[] { '=' }, 2);

            if (parts.Length != 2)
            {
                continue;
            }

            string key = parts[0].Trim();
            string value = parts[1].Trim();

            switch (key)
            {
                case "FolderActionCommand":
                    config.FolderActionCommand = value;
                    break;
                case "AppTheme":
                    Theme parsedTheme;
                    if (Enum.TryParse(value, out parsedTheme))
                    {
                        config.AppTheme = parsedTheme;
                    }
                    break;
                case "MenuFontSize":
                    FontSize parsedSize;
                    if (Enum.TryParse(value, out parsedSize))
                    {
                        config.MenuFontSize = parsedSize;
                    }
                    break;
            }
        }

        return config;
    }

    public void Save()
    {
        List<string> lines = new List<string>
        {
            "FolderActionCommand=" + (FolderActionCommand ?? ""),
            "AppTheme=" + AppTheme.ToString(),
            "MenuFontSize=" + MenuFontSize.ToString()
        };

        File.WriteAllLines(ConfigFilePath, lines);
    }
}
