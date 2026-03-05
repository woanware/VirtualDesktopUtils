using System.Diagnostics;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace VirtualDesktopUtils.Services;

internal enum AppUpdateCheckStatus
{
    UpToDate,
    UpdateAvailable,
    Error
}

internal sealed record AppReleaseInfo(
    string VersionText,
    Version Version,
    string ReleaseUrl,
    string ReleaseTitle,
    string ReleaseNotes);

internal sealed record AppUpdateCheckResult(
    AppUpdateCheckStatus Status,
    string Message,
    Version CurrentVersion,
    AppReleaseInfo? LatestRelease);

internal sealed class AppUpdateService
{
    private const string ReleasesApiUrl = "https://api.github.com/repos/woanware/VirtualDesktopUtils/releases";
    private static readonly Regex VersionTagPattern = new(
        "^v?(?<version>\\d+\\.\\d+\\.\\d+(?:\\.\\d+)?)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HttpClient HttpClient = CreateHttpClient();

    public async Task<AppUpdateCheckResult> CheckForUpdatesAsync(CancellationToken cancellationToken = default)
    {
        var currentVersion = GetCurrentVersion();
        List<GitHubReleaseDto>? releaseDtos;

        try
        {
            using var response = await HttpClient.GetAsync(ReleasesApiUrl, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return new AppUpdateCheckResult(
                    AppUpdateCheckStatus.Error,
                    $"Update check failed: GitHub returned {(int)response.StatusCode}.",
                    currentVersion,
                    null);
            }

            await using var contentStream = await response.Content.ReadAsStreamAsync(cancellationToken);
            releaseDtos = await JsonSerializer.DeserializeAsync<List<GitHubReleaseDto>>(contentStream, cancellationToken: cancellationToken);
        }
        catch (HttpRequestException ex)
        {
            return new AppUpdateCheckResult(AppUpdateCheckStatus.Error, $"Update check failed: {ex.Message}", currentVersion, null);
        }
        catch (TaskCanceledException)
        {
            return new AppUpdateCheckResult(AppUpdateCheckStatus.Error, "Update check failed: request timed out.", currentVersion, null);
        }
        catch (JsonException ex)
        {
            return new AppUpdateCheckResult(AppUpdateCheckStatus.Error, $"Update check failed: invalid response ({ex.Message}).", currentVersion, null);
        }

        if (releaseDtos is null || releaseDtos.Count == 0)
        {
            return new AppUpdateCheckResult(AppUpdateCheckStatus.Error, "Update check failed: no releases found.", currentVersion, null);
        }

        var latestRelease = releaseDtos
            .Where(release => !release.Draft)
            .Where(release => !release.Prerelease)
            .Select(TryCreateReleaseInfo)
            .OfType<AppReleaseInfo>()
            .OrderByDescending(release => NormalizeVersion(release.Version))
            .FirstOrDefault();

        if (latestRelease is null)
        {
            return new AppUpdateCheckResult(
                AppUpdateCheckStatus.Error,
                "Update check failed: no valid release versions found.",
                currentVersion,
                null);
        }

        if (NormalizeVersion(latestRelease.Version) <= NormalizeVersion(currentVersion))
        {
            return new AppUpdateCheckResult(
                AppUpdateCheckStatus.UpToDate,
                $"You are up to date (v{FormatVersion(currentVersion)}).",
                currentVersion,
                latestRelease);
        }

        return new AppUpdateCheckResult(
            AppUpdateCheckStatus.UpdateAvailable,
            $"Update available: v{latestRelease.VersionText} (current: v{FormatVersion(currentVersion)}).",
            currentVersion,
            latestRelease);
    }

    public static Version GetCurrentVersion()
    {
        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath))
        {
            try
            {
                var fileVersion = FileVersionInfo.GetVersionInfo(processPath).FileVersion;
                if (!string.IsNullOrWhiteSpace(fileVersion) && Version.TryParse(fileVersion, out var parsedFileVersion))
                {
                    return parsedFileVersion;
                }
            }
            catch (ArgumentException)
            {
            }
            catch (System.IO.IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        var assemblyVersion = typeof(AppUpdateService).Assembly.GetName().Version;
        return assemblyVersion ?? new Version(0, 0, 0, 0);
    }

    private static AppReleaseInfo? TryCreateReleaseInfo(GitHubReleaseDto release)
    {
        if (string.IsNullOrWhiteSpace(release.TagName) || string.IsNullOrWhiteSpace(release.HtmlUrl))
        {
            return null;
        }

        var match = VersionTagPattern.Match(release.TagName.Trim());
        if (!match.Success)
        {
            return null;
        }

        var versionText = match.Groups["version"].Value;
        if (!Version.TryParse(versionText, out var version))
        {
            return null;
        }

        var title = !string.IsNullOrWhiteSpace(release.Name) ? release.Name : release.TagName;
        var notes = release.Body ?? string.Empty;
        return new AppReleaseInfo(versionText, version, release.HtmlUrl, title, notes);
    }

    private static Version NormalizeVersion(Version version) =>
        new(version.Major, version.Minor, version.Build < 0 ? 0 : version.Build, version.Revision < 0 ? 0 : version.Revision);

    private static string FormatVersion(Version version)
    {
        var normalized = NormalizeVersion(version);
        return normalized.Revision == 0
            ? $"{normalized.Major}.{normalized.Minor}.{normalized.Build}"
            : normalized.ToString();
    }

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(12)
        };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("VirtualDesktopUtils-Updater/1.0");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    private sealed class GitHubReleaseDto
    {
        [JsonPropertyName("tag_name")]
        public string? TagName { get; init; }

        [JsonPropertyName("name")]
        public string? Name { get; init; }

        [JsonPropertyName("html_url")]
        public string? HtmlUrl { get; init; }

        [JsonPropertyName("body")]
        public string? Body { get; init; }

        [JsonPropertyName("draft")]
        public bool Draft { get; init; }

        [JsonPropertyName("prerelease")]
        public bool Prerelease { get; init; }
    }
}
