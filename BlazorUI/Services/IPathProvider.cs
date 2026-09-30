namespace BlazorUI.Services;

public interface IPathProvider
{
    string? SelectedBank { get; }
    string CurrentPath { get; }
    string? ServerPath { get; }
    void SelectBank(string bank);
}
