using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

    private async Task CheckForUpdatesOnStartupAsync()
    {
        try
        {
            latestRelease = await GetLatestReleaseAsync();
        }
        catch
        {
            latestRelease = null;
        }
    }

    private async Task PromptForUpdateIfAvailableAsync()
    {
        if (!IsUpdateAvailable() || latestRelease is null)
        {
            return;
        }

        AnsiConsole.Clear();
        RenderTopNav("Home");
        AnsiConsole.Write(new Panel(new Rows(
                new Markup($"[bold yellow]Update available[/]"),
                new Markup($"Current version: [grey]v{Markup.Escape(GetCurrentVersion())}[/]"),
                new Markup($"Latest release: [grey]{Markup.Escape(latestRelease.TagName)}[/]"),
                new Markup($"Release URL: [grey]{Markup.Escape(latestRelease.HtmlUrl)}[/]")))
            .Header("W365 CLI")
            .Border(BoxBorder.Rounded)
            .BorderStyle(new Style(Color.Yellow)));

        var installNow = AskYesNo("Download and install this update now?");
        if (!installNow)
        {
            var openRelease = AskYesNo("Open the latest GitHub release page instead?", defaultToYes: false);
            if (openRelease)
            {
                OpenUrl(latestRelease.HtmlUrl);
                TimedMessage("[green]Opened latest release.[/]", 1200);
            }

            return;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            await DownloadAndInstallWindowsUpdateAsync(latestRelease);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            await DownloadAndInstallMacUpdateAsync(latestRelease);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            await DownloadAndInstallLinuxUpdateAsync(latestRelease);
        }
        else
        {
            AnsiConsole.MarkupLine("[yellow]Automatic updates aren't supported on this platform yet. Opening the release page instead.[/]");
            OpenUrl(latestRelease.HtmlUrl);
            TimedMessage("[green]Opened latest release.[/]", 1500);
        }
    }

    private static string GetCurrentOsArch() => RuntimeInformation.ProcessArchitecture switch
    {
        Architecture.Arm64 => "arm64",
        _ => "x64"
    };

    private static GitHubReleaseAsset? FindWindowsInstallerAsset(GitHubReleaseInfo release)
    {
        var suffix = $"win-{GetCurrentOsArch()}.exe";
        return release.Assets.FirstOrDefault(a =>
            a.Name.StartsWith("W365CLISetup-", StringComparison.OrdinalIgnoreCase) &&
            a.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
    }

    private static GitHubReleaseAsset? FindMacZipAsset(GitHubReleaseInfo release)
    {
        var name = $"w365-osx-{GetCurrentOsArch()}.zip";
        return release.Assets.FirstOrDefault(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private static GitHubReleaseAsset? FindLinuxTarAsset(GitHubReleaseInfo release)
    {
        var name = $"w365-linux-{GetCurrentOsArch()}.tar.gz";
        return release.Assets.FirstOrDefault(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private static async Task DownloadFileAsync(string url, string destinationPath)
    {
        using var http = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5)
        };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("W365CliNative");
        await using var responseStream = await http.GetStreamAsync(url);
        await using var fileStream = File.Create(destinationPath);
        await responseStream.CopyToAsync(fileStream);
    }

    private static async Task DownloadAndInstallWindowsUpdateAsync(GitHubReleaseInfo release)
    {
        var asset = FindWindowsInstallerAsset(release);
        if (asset is null)
        {
            AnsiConsole.MarkupLine($"[yellow]Couldn't find a Windows installer for this release ({Markup.Escape(GetCurrentOsArch())}). Opening the release page instead.[/]");
            OpenUrl(release.HtmlUrl);
            WaitForAnyKey();
            return;
        }

        var tempPath = Path.Combine(Path.GetTempPath(), asset.Name);
        try
        {
            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Downloading {asset.Name}...", async _ => await DownloadFileAsync(asset.BrowserDownloadUrl, tempPath));
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Download failed:[/] [grey]{Markup.Escape(ex.Message)}[/]");
            var retry = AskYesNo("This is often a transient network hiccup. Try downloading again?");
            if (retry)
            {
                await DownloadAndInstallWindowsUpdateAsync(release);
                return;
            }

            AnsiConsole.MarkupLine("[grey]Opening the release page so you can download it manually.[/]");
            OpenUrl(release.HtmlUrl);
            WaitForAnyKey();
            return;
        }

        TimedMessage($"[green]Downloaded {Markup.Escape(asset.Name)}.[/]", 1000);

        var runNow = AskYesNo("Run the installer now? W365 CLI will close so it can update — it installs silently and only takes a few seconds.");
        if (!runNow)
        {
            AnsiConsole.MarkupLine($"[grey]Saved the installer to:[/] [white]{Markup.Escape(tempPath)}[/]");
            AnsiConsole.MarkupLine("[grey]Double-click it anytime to update.[/]");
            var reveal = AskYesNo("Open the folder containing the installer?", defaultToYes: false);
            if (reveal)
            {
                try { Process.Start("explorer.exe", $"/select,\"{tempPath}\""); } catch { /* best effort */ }
            }

            WaitForAnyKey("[grey]Press any key to continue — you can update whenever you're ready.[/]");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(tempPath, "/VERYSILENT /NORESTART")
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Couldn't launch the installer:[/] [grey]{Markup.Escape(ex.Message)}[/]");
            AnsiConsole.MarkupLine($"[grey]You can run it manually from:[/] [white]{Markup.Escape(tempPath)}[/]");
            WaitForAnyKey();
            return;
        }

        AnsiConsole.MarkupLine("[green]Installer launched.[/] [grey]W365 CLI is closing so it can finish updating — reopen it in a few seconds.[/]");
        Thread.Sleep(1200);
        Environment.Exit(0);
    }

    /// <summary>
    /// Shared download → extract → atomic-replace flow for macOS and Linux self-updates — the two
    /// platforms only differ in archive format (zip vs tar.gz) and one macOS-only quarantine-flag
    /// cleanup step, so this holds the common retry/error/messaging logic once instead of twice.
    /// </summary>
    [System.Runtime.Versioning.SupportedOSPlatform("macos")]
    [System.Runtime.Versioning.SupportedOSPlatform("linux")]
    private static async Task DownloadAndInstallUnixUpdateAsync(
        GitHubReleaseInfo release,
        GitHubReleaseAsset? asset,
        string platformName,
        Action<string, string> extractArchive,
        Action<string>? afterExtract = null)
    {
        if (asset is null)
        {
            AnsiConsole.MarkupLine($"[yellow]Couldn't find a {Markup.Escape(platformName)} build for this release ({Markup.Escape(GetCurrentOsArch())}). Opening the release page instead.[/]");
            OpenUrl(release.HtmlUrl);
            WaitForAnyKey();
            return;
        }

        var processPath = Environment.ProcessPath;
        var canReplaceInPlace = processPath is not null &&
            !string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase);

        var tempDir = Path.Combine(Path.GetTempPath(), "w365cli-update-" + Guid.NewGuid().ToString("N"));
        var cleanUpTempDir = true;
        try
        {
            Directory.CreateDirectory(tempDir);
            var archivePath = Path.Combine(tempDir, asset.Name);

            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Downloading {asset.Name}...", async _ => await DownloadFileAsync(asset.BrowserDownloadUrl, archivePath));

            var extractDir = Path.Combine(tempDir, "extracted");
            Directory.CreateDirectory(extractDir);
            extractArchive(archivePath, extractDir);

            var newBinary = Directory.GetFiles(extractDir, "W365Cli", SearchOption.AllDirectories).FirstOrDefault();
            if (newBinary is null)
            {
                AnsiConsole.MarkupLine("[red]Couldn't find the W365Cli binary inside the downloaded archive.[/]");
                AnsiConsole.MarkupLine("[grey]Opening the release page so you can update manually.[/]");
                OpenUrl(release.HtmlUrl);
                WaitForAnyKey();
                return;
            }

            const UnixFileMode ExecutablePermissions =
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
                UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                UnixFileMode.OtherRead | UnixFileMode.OtherExecute;
            File.SetUnixFileMode(newBinary, ExecutablePermissions);

            afterExtract?.Invoke(newBinary);

            if (!canReplaceInPlace || processPath is null)
            {
                cleanUpTempDir = false;
                AnsiConsole.MarkupLine("[yellow]Downloaded the update, but couldn't determine where W365 CLI is installed to replace it automatically.[/]");
                AnsiConsole.MarkupLine($"[grey]New binary saved to:[/] [white]{Markup.Escape(newBinary)}[/]");
                AnsiConsole.MarkupLine("[grey]Copy it over your installed w365cli binary (commonly ~/.local/bin/w365cli) to finish updating.[/]");
                WaitForAnyKey();
                return;
            }

            // Atomic replace: copy the new binary next to the running one, then rename over it.
            // rename() on Unix swaps the directory entry without touching the inode the currently
            // running process still has open, so this is safe even while w365cli is executing.
            var targetDir = Path.GetDirectoryName(processPath)!;
            var stagingPath = Path.Combine(targetDir, ".w365cli.update.tmp");
            File.Copy(newBinary, stagingPath, overwrite: true);
            File.SetUnixFileMode(stagingPath, ExecutablePermissions);
            File.Move(stagingPath, processPath, overwrite: true);

            // Deliberately not exiting/relaunching here. Forcing this process to exit (via
            // Environment.Exit or by spawning a replacement and killing this one) can leave the
            // terminal's tty/termios settings in a bad state — since Console.ReadKey's raw-mode
            // handling doesn't get a chance to clean up on an abrupt exit — which then makes the
            // *next* process launched in that same terminal window crash with an Input/output
            // error the moment it tries to read a key (observed in practice on macOS; the same
            // termios mechanism applies on Linux terminals too). The binary on disk is already
            // updated; this session just keeps running the old in-memory build until the user
            // exits normally and reopens w365cli themselves.
            try { Directory.Delete(tempDir, recursive: true); } catch { /* best effort cleanup */ }
            cleanUpTempDir = false;
            AnsiConsole.MarkupLine($"[green]Updated to {Markup.Escape(release.TagName)}.[/]");
            AnsiConsole.MarkupLine("[grey]The new version will be used the next time you quit and reopen w365cli.[/]");
            WaitForAnyKey();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Update failed:[/] [grey]{Markup.Escape(ex.Message)}[/]");
            var retry = AskYesNo("This is often a transient network hiccup. Try again?");
            if (retry)
            {
                if (cleanUpTempDir)
                {
                    try { Directory.Delete(tempDir, recursive: true); } catch { /* best effort cleanup */ }
                    cleanUpTempDir = false;
                }

                await DownloadAndInstallUnixUpdateAsync(release, asset, platformName, extractArchive, afterExtract);
                return;
            }

            AnsiConsole.MarkupLine("[grey]Opening the release page so you can update manually.[/]");
            OpenUrl(release.HtmlUrl);
            WaitForAnyKey();
        }
        finally
        {
            if (cleanUpTempDir)
            {
                try { Directory.Delete(tempDir, recursive: true); } catch { /* best effort cleanup */ }
            }
        }
    }

    [System.Runtime.Versioning.SupportedOSPlatform("macos")]
    private static Task DownloadAndInstallMacUpdateAsync(GitHubReleaseInfo release)
    {
        var asset = FindMacZipAsset(release);
        return DownloadAndInstallUnixUpdateAsync(
            release,
            asset,
            "macOS",
            extractArchive: (archivePath, extractDir) => System.IO.Compression.ZipFile.ExtractToDirectory(archivePath, extractDir),
            afterExtract: newBinary =>
            {
                // Best-effort: clear the quarantine flag in case this ever gets flagged (matches install.sh).
                try
                {
                    var xattrProcess = Process.Start(new ProcessStartInfo("xattr", $"-d com.apple.quarantine \"{newBinary}\"")
                    {
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    });
                    xattrProcess?.WaitForExit(2000);
                }
                catch { /* best effort */ }
            });
    }

    [System.Runtime.Versioning.SupportedOSPlatform("linux")]
    private static Task DownloadAndInstallLinuxUpdateAsync(GitHubReleaseInfo release)
    {
        var asset = FindLinuxTarAsset(release);
        return DownloadAndInstallUnixUpdateAsync(
            release,
            asset,
            "Linux",
            extractArchive: ExtractTarGz);
    }

    /// <summary>
    /// TarFile.ExtractToDirectory(string sourceFileName, ...) reads the file as a PLAIN tar
    /// stream -- it does NOT auto-decompress gzip despite happily accepting a .tar.gz path with no
    /// error at the call site. Feeding it our gzip-compressed release archive directly meant it was
    /// parsing raw gzip magic bytes as if they were tar header fields, which is why every Linux
    /// self-update failed with a misleading "InvalidDataException: Unable to parse number" caught
    /// by the generic handler and mislabeled as a "transient network hiccup" -- the download itself
    /// was always fine. Manually decompressing via GZipStream first and handing THAT stream to the
    /// Stream-accepting ExtractToDirectory overload is what actually works. Verified against a real
    /// downloaded release asset: the string-path overload threw on both plain, ustar, and pax tar
    /// variants of the same gzip wrapper, while gunzip-then-extract succeeded immediately.
    /// </summary>
    [System.Runtime.Versioning.SupportedOSPlatform("linux")]
    private static void ExtractTarGz(string archivePath, string extractDir)
    {
        using var fileStream = File.OpenRead(archivePath);
        using var gzipStream = new System.IO.Compression.GZipStream(fileStream, System.IO.Compression.CompressionMode.Decompress);
        System.Formats.Tar.TarFile.ExtractToDirectory(gzipStream, extractDir, overwriteFiles: true);
    }

    private bool IsUpdateAvailable()
    {
        if (latestRelease is null)
        {
            return false;
        }

        var current = ParseVersion(GetCurrentVersion());
        var latest = ParseVersion(latestRelease.TagName);
        return latest is not null && current is not null && latest > current;
    }

    private static async Task<GitHubReleaseInfo> GetLatestReleaseAsync()
    {
        using var http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(5)
        };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("W365CliNative");
        await using var stream = await http.GetStreamAsync(GitHubLatestReleaseApiUrl);
        using var document = await JsonDocument.ParseAsync(stream);
        var root = document.RootElement;
        var assets = new List<GitHubReleaseAsset>();
        if (root.TryGetProperty("assets", out var assetsElement) && assetsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var assetElement in assetsElement.EnumerateArray())
            {
                var name = assetElement.TryGetProperty("name", out var nameEl) ? nameEl.GetString() : null;
                var assetUrl = assetElement.TryGetProperty("browser_download_url", out var urlEl) ? urlEl.GetString() : null;
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(assetUrl))
                {
                    assets.Add(new GitHubReleaseAsset(name, assetUrl));
                }
            }
        }

        return new GitHubReleaseInfo(
            root.TryGetProperty("tag_name", out var tag) ? tag.GetString() ?? "unknown" : "unknown",
            root.TryGetProperty("html_url", out var url) ? url.GetString() ?? GitHubRepositoryUrl : GitHubRepositoryUrl,
            root.TryGetProperty("published_at", out var published) && published.ValueKind == JsonValueKind.String && DateTimeOffset.TryParse(published.GetString(), out var publishedAt)
                ? publishedAt
                : null,
            assets);
    }

    private sealed record GitHubReleaseInfo(string TagName, string HtmlUrl, DateTimeOffset? PublishedAt, IReadOnlyList<GitHubReleaseAsset> Assets);

    private sealed record GitHubReleaseAsset(string Name, string BrowserDownloadUrl);
}
