namespace ExcelForge;

internal static class AppConstants
{

    public const string AppName        = "ExcelForge";
    public const string Version        = "0.2.0";
    public const string VersionDisplay = "v0.2.0";

    public const string GitHubOwner   = "BaadHackor";
    public const string GitHubRepo    = "excelforge";
    public const string GitHubApiUrl  =
        $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";

    public const string UpdaterUserAgent =
        $"ExcelForge/{Version} (auto-updater; github-releases; win-x64)";

    public const string NativeAssetName = "ExcelForge.Native.dll";

    public const string NativeDllName  = "ExcelForge.Native.dll";

    public const string EventLogSource = "ExcelForge";
    public const string EventLogName   = "Application";

    public const int    MaxRecentFiles = 10;
    public const string WebsiteUrl     = "https://github.com/BaadHackor/excelforge";
}
