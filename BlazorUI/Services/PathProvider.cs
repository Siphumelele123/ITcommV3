using BlazorUI.Models;
using Microsoft.Extensions.Options;

namespace BlazorUI.Services;

public class PathProvider : IPathProvider
{
    private readonly FileSettings _settings;

    public PathProvider(IOptions<FileSettings> options)
    {
        _settings = options.Value;
        CurrentPath = _settings.DefaultPath;
    }

    public string? SelectedBank { get; private set; }
    public string CurrentPath { get; private set; }
    public string? ServerPath { get; private set; }

    public void SelectBank(string bank)
    {
        if (!_settings.Banks.TryGetValue(bank, out var paths))
            throw new ArgumentException($"Unknown bank selection: {bank}");

        SelectedBank = bank;
        CurrentPath = paths.DesktopFolder;
        ServerPath = paths.ServerPath;
    }
}
