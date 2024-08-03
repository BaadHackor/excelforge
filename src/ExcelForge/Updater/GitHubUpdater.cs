using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace ExcelForge.Updater;

internal sealed class GitHubUpdater : IDisposable
{
    private static readonly HttpClient _http = new(new HttpClientHandler
    {
        AllowAutoRedirect        = true,
        MaxAutomaticRedirections = 5,
    });

    private readonly ILogger<GitHubUpdater> _logger;
    private bool _disposed;

    static GitHubUpdater()
    {
        _http.Timeout = TimeSpan.FromSeconds(15);

        _http.DefaultRequestHeaders.UserAgent.ParseAdd(AppConstants.UpdaterUserAgent);
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        _http.DefaultRequestHeaders.Add("X-GitHub-Api-Version", "2022-11-28");
    }

    public GitHubUpdater(ILogger<GitHubUpdater> logger)
    {
        _logger = logger;
    }

    public async Task<bool> CheckAndApplyAsync(bool printToConsole = false)
    {
        _logger.LogInformation("Updater: querying {Url}", AppConstants.GitHubApiUrl);

        ReleaseInfo? release;
        try
        {
            release = await FetchLatestReleaseAsync();
        }
        catch (HttpRequestException ex)
        {
            _logger.LogWarning("Updater: network error — {Msg}", ex.Message);
            return false;
        }
        catch (TaskCanceledException)
        {
            _logger.LogWarning("Updater: request timed out");
            return false;
        }

        if (release is null || release.Draft || release.Prerelease)
        {
            _logger.LogInformation("Updater: no stable release found");
            return false;
        }

        if (!IsNewer(release.SemanticVersion, AppConstants.Version))
        {
            _logger.LogInformation(
                "Updater: already on latest ({Current})", AppConstants.Version);
            if (printToConsole)
                Console.WriteLine($"Already up to date ({AppConstants.VersionDisplay}).");
            return false;
        }

        _logger.LogInformation(
            "Updater: new release found — {Tag} (current: {Current})",
            release.TagName, AppConstants.Version);

        if (printToConsole)
            Console.WriteLine($"Update available: {release.TagName}  ({release.PublishedAt:yyyy-MM-dd})");

        var nativeAsset = release.Assets
            .FirstOrDefault(a => a.Name.Equals(
                AppConstants.NativeAssetName, StringComparison.OrdinalIgnoreCase));

        if (nativeAsset is null)
        {
            _logger.LogWarning(
                "Updater: release {Tag} has no asset named '{Asset}' — skipping",
                release.TagName, AppConstants.NativeAssetName);
            return false;
        }

        try
        {
            await DownloadAndReplaceNativeDllAsync(nativeAsset, release.TagName);
            LogUpdateEvent(release.TagName);
            if (printToConsole)
                Console.WriteLine($"ExcelForge.Native.dll updated to {release.TagName}.");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Updater: failed to apply update from {Tag}", release.TagName);
            return false;
        }
    }

    private async Task<ReleaseInfo?> FetchLatestReleaseAsync()
    {
        string json = await _http.GetStringAsync(AppConstants.GitHubApiUrl);

        _logger.LogDebug("Updater: received {Bytes} bytes from GitHub API", json.Length);

        return JsonSerializer.Deserialize<ReleaseInfo>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
        });
    }

    private async Task DownloadAndReplaceNativeDllAsync(ReleaseAsset asset, string tag)
    {
        string appDir  = AppDomain.CurrentDomain.BaseDirectory;
        string dllPath = Path.Combine(appDir, AppConstants.NativeDllName);
        string tmpPath = dllPath + ".tmp";

        _logger.LogInformation(
            "Updater: downloading {Name} ({Size:N0} bytes) from {Url}",
            asset.Name, asset.Size, asset.DownloadUrl);

        byte[] data = await _http.GetByteArrayAsync(asset.DownloadUrl);

        _logger.LogInformation("Updater: download complete ({Bytes:N0} bytes received)", data.Length);

        await File.WriteAllBytesAsync(tmpPath, data);

        File.Move(tmpPath, dllPath, overwrite: true);

        _logger.LogInformation(
            "Updater: {Dll} replaced successfully (release {Tag})",
            AppConstants.NativeDllName, tag);

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(500);
                LoadLibrary(dllPath);
            }
            catch { }
        });
    }

    private void LogUpdateEvent(string newVersion)
    {
        try
        {
            if (!EventLog.SourceExists(AppConstants.EventLogSource))
                EventLog.CreateEventSource(AppConstants.EventLogSource, AppConstants.EventLogName);

            using var evLog = new EventLog(AppConstants.EventLogName)
            {
                Source = AppConstants.EventLogSource,
            };

            evLog.WriteEntry(
                $"{AppConstants.AppName} component updated ({newVersion}).",
                EventLogEntryType.Information,
                eventID: 1000);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Updater: could not write to Windows Event Log");
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibrary(string lpFileName);

    private static bool IsNewer(string remote, string local)
    {
        if (!Version.TryParse(remote, out var r)) return false;
        if (!Version.TryParse(local,  out var l)) return false;
        return r > l;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
    }
}
