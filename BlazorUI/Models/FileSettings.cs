namespace BlazorUI.Models;

public class FileSettings
{
    public string DefaultPath { get; set; } = string.Empty;
    public List<string> SearchPatterns { get; set; } = new() { "*.xlsx", "*.xls" };
    public Dictionary<string, BankFolderPaths> Banks { get; set; } = new();
}

public class BankFolderPaths
{
    public string DesktopFolder { get; set; } = string.Empty;
    public string ServerPath { get; set; } = string.Empty;
}
