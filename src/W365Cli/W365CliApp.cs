using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed class W365CliApp
{
    private const string GitHubRepositoryUrl = "https://github.com/bwya77/W365-CLI-Native";
    private const string GitHubLatestReleaseApiUrl = "https://api.github.com/repos/bwya77/W365-CLI-Native/releases/latest";
    private const string ProjectWebsiteUrl = "https://www.windowsfromanywhere.com";
    private const string GitHubFeatureUrl = "https://github.com/bwya77/W365-CLI-Native/issues/new?labels=enhancement&title=Feature%3A%20";
    private const string GitHubIssueUrl = "https://github.com/bwya77/W365-CLI-Native/issues/new?labels=bug&title=Bug%3A%20";
    private const string AccentColor = "#58a6ff";
    private const string AccentSoftColor = "#79c0ff";
    private const string MutedColor = "#8b949e";
    private const string TextColor = "#f0f6fc";
    private const string SurfaceColor = "#161b22";
    private const string PurpleColor = "#bc8cff";
    private const string GreenColor = "#3fb950";
    private const string BorderColor = "#30363d";
    private readonly W365Session _session = new();
    private readonly Tip launchTip = Tips[Random.Shared.Next(Tips.Length)];
    private static readonly List<ActionHistoryItem> ActionHistory = [];
    private static readonly Tip[] Tips =
    [
        new("/palette", "Use P or Ctrl+K to open the command palette from anywhere on the main screen."),
        new("/filter", "Use slash or F on list screens to filter down to the thing you need."),
        new("/refresh", "Press R on data-heavy screens to refresh without leaving your current view."),
        new("/history", "Press H to open hidden action history for submitted Cloud PC actions."),
        new("/resize", "Resize uses an interactive service plan picker so you can compare sizes before submitting."),
        new("/licensing", "Licensing shows Flex, Reserve, Cloud Apps, shared pool, and dedicated capacity together."),
        new("/cloudapps", "Cloud Apps can be browsed, published, and unpublished from the native CLI."),
        new("/snapshots", "Snapshots can be reviewed across all Cloud PCs or opened from a specific Cloud PC."),
        new("/reports", "Report rows can open Cloud PC detail pages when Graph returns enough identifying data."),
        new("/safe-actions", "Device-impacting actions use confirmation screens before submission.")
    ];
    private static string? statusMessage;
    private static DateTimeOffset? statusMessageAt;
    private static string statusBarConnection = "[white on red] NOT CONNECTED [/]";
    private static string statusBarTenant = "No tenant selected";
    private GitHubReleaseInfo? latestRelease;

    public async Task<int> RunAsync(string[] args)
    {
        // Console.Title's setter throws PlatformNotSupportedException on non-Windows platforms —
        // this app ships macOS builds too, so guard it instead of crashing at startup.
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Console.Title = "W365 CLI";
        }

        Console.CursorVisible = false;
        Console.CancelKeyPress += (_, _) => Console.CursorVisible = true;
        await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Checking cached sign-in...", async _ => await _session.TryRestoreAsync());
        await ShowMissingPermissionPromptIfNeededAsync();
        await CheckForUpdatesOnStartupAsync();
        await PromptForUpdateIfAvailableAsync();

        var selectedIndex = 0;
        var expandedIndex = -1;
        var selectedChildIndex = -1;
        var topNavIndex = -1;
        while (true)
        {
            try
            {
                var menuChoices = GetMainMenuChoices();
                if (selectedIndex >= menuChoices.Count)
                {
                    selectedIndex = menuChoices.Count - 1;
                    expandedIndex = -1;
                    selectedChildIndex = -1;
                }

                RenderMainMenu(menuChoices, selectedIndex, expandedIndex, selectedChildIndex, topNavIndex);
                var key = ReadNavigationKey(intercept: true, handleTopNavTab: false);
                var selectedChoice = menuChoices[selectedIndex];
                var selectedChildren = selectedChoice.Children ?? [];

                if (TryHandleTopNavKey(key, ref topNavIndex, currentTabIndex: 0, out var activation))
                {
                    switch (activation)
                    {
                        case TopNavActivation.Home:
                            selectedIndex = 0;
                            expandedIndex = -1;
                            selectedChildIndex = -1;
                            break;
                        case TopNavActivation.About:
                            try
                            {
                                ShowAbout();
                            }
                            catch (NavigateHomeException)
                            {
                                selectedIndex = 0;
                                expandedIndex = -1;
                                selectedChildIndex = -1;
                                topNavIndex = -1;
                            }
                            catch (NavigateExitException)
                            {
                                AnsiConsole.Clear();
                                Console.CursorVisible = true;
                                return 0;
                            }
                            break;
                        case TopNavActivation.Exit:
                            AnsiConsole.Clear();
                            Console.CursorVisible = true;
                            return 0;
                    }
                    continue;
                }

                switch (key.Key)
                {
                    case ConsoleKey.UpArrow:
                        if (expandedIndex == selectedIndex && selectedChildIndex >= 0)
                        {
                            selectedChildIndex--;
                        }
                        else
                        {
                            selectedIndex = Math.Max(0, selectedIndex - 1);
                            expandedIndex = -1;
                            selectedChildIndex = -1;
                        }
                        break;
                    case ConsoleKey.DownArrow:
                        if (expandedIndex == selectedIndex && selectedChildren.Count > 0 && selectedChildIndex < selectedChildren.Count - 1)
                        {
                            selectedChildIndex++;
                        }
                        else
                        {
                            selectedIndex = Math.Min(menuChoices.Count - 1, selectedIndex + 1);
                            expandedIndex = -1;
                            selectedChildIndex = -1;
                        }
                        break;
                    case ConsoleKey.Home:
                        selectedIndex = 0;
                        expandedIndex = -1;
                        selectedChildIndex = -1;
                        break;
                    case ConsoleKey.End:
                        selectedIndex = menuChoices.Count - 1;
                        expandedIndex = -1;
                        selectedChildIndex = -1;
                        break;
                    case ConsoleKey.RightArrow:
                        if (selectedChildren.Count > 0)
                        {
                            expandedIndex = selectedIndex;
                            selectedChildIndex = selectedChildIndex < 0 ? 0 : Math.Min(selectedChildren.Count - 1, selectedChildIndex + 1);
                        }
                        break;
                    case ConsoleKey.Enter:
                        if (selectedChildren.Count > 0 && selectedChildIndex < 0)
                        {
                            expandedIndex = selectedIndex;
                            selectedChildIndex = 0;
                            break;
                        }

                        var choiceToExecute = selectedChildIndex >= 0
                            ? selectedChildren[selectedChildIndex]
                            : selectedChoice;
                        if (await ExecuteMainMenuChoiceAsync(choiceToExecute))
                        {
                            return 0;
                        }
                        break;
                    case ConsoleKey.Escape:
                    case ConsoleKey.LeftArrow:
                        if (expandedIndex >= 0)
                        {
                            expandedIndex = -1;
                            selectedChildIndex = -1;
                        }
                        break;
                    case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                        await ShowCommandPaletteAsync();
                        break;
                    default:
                        if (IsActionHistoryHotkey(key))
                        {
                            await ShowActionHistoryAsync();
                        }
                        else if (key.KeyChar is 'p' or 'P')
                        {
                            await ShowCommandPaletteAsync();
                        }
                        else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                        {
                            expandedIndex = -1;
                            selectedChildIndex = -1;
                        }
                        break;
                }
            }
            catch (NavigateHomeException)
            {
                selectedIndex = 0;
                expandedIndex = -1;
                selectedChildIndex = -1;
                topNavIndex = -1;
            }
            catch (NavigateAboutException)
            {
                try
                {
                    ShowAbout();
                }
                catch (NavigateHomeException)
                {
                    selectedIndex = 0;
                    expandedIndex = -1;
                    selectedChildIndex = -1;
                    topNavIndex = -1;
                }
                catch (NavigateExitException)
                {
                    AnsiConsole.Clear();
                    Console.CursorVisible = true;
                    return 0;
                }
            }
            catch (NavigateExitException)
            {
                AnsiConsole.Clear();
                Console.CursorVisible = true;
                return 0;
            }
        }
    }

    private async Task<bool> ExecuteMainMenuChoiceAsync(MenuChoice choice)
    {
        switch (choice.Title)
        {
            case "Browse Cloud PCs":
                await ShowCloudPcsAsync();
                return false;
            case "Disk space":
                await ShowDiskSpaceAsync();
                return false;
            case "Snapshots":
                await ShowAllSnapshotsAsync();
                return false;
            case "Policies":
                await ShowProvisioningAsync();
                return false;
            case "Create policy":
                await CreateProvisioningPolicyWizardAsync();
                return false;
            case "User experience sync overview":
                await ShowUserExperienceSyncOverviewAsync();
                return false;
            case "Usage report":
            case "Usage":
            case "Sign-in status":
                await ShowGraphRowsAsync(
                    "Windows 365 Cloud PC sign-in status",
                    async () => await _session.Graph.GetSignInStatusRowsAsync(),
                    GetUsageReportHeader,
                    FormatUsageReportRow,
                    OpenCloudPcFromReportRowAsync);
                return false;
            case "Connectivity history":
                await ShowConnectivityHistoryAsync();
                return false;
            case "Launch details":
                await ShowGraphRowsAsync(
                    "Windows 365 launch details",
                    async () => await _session.Graph.GetLaunchDetailRowsAsync(),
                    GetLaunchDetailsHeader,
                    FormatLaunchDetailsRow);
                return false;
            case "Cloud PC reports":
                await ShowCloudPcReportsAsync();
                return false;
            case "Organization settings":
                await ShowGraphRowsAsync(
                    "Windows 365 organization settings",
                    async () => await _session.Graph.GetOrganizationSettingsAsync(),
                    GetOrganizationSettingsHeader,
                    FormatOrganizationSettingRow);
                return false;
            case "Setting profiles":
                await ShowGraphRowsAsync(
                    "Windows 365 setting profiles",
                    async () => await _session.Graph.GetSettingProfilesAsync(),
                    GetSettingProfilesHeader,
                    FormatSettingProfileRow);
                return false;
            case "User settings":
                await ShowGraphRowsAsync(
                    "Windows 365 user settings",
                    async () => await _session.Graph.GetUserSettingsAsync(),
                    GetUserSettingsHeader,
                    FormatUserSettingRow);
                return false;
            case "Service plans":
                await ShowGraphRowsAsync("Windows 365 service plans", _session.Graph.GetServicePlanRowsAsync, GetServicePlansHeader, FormatServicePlanRow);
                return false;
            case "Gallery images":
                await ShowGraphRowsAsync("Windows 365 gallery images", _session.Graph.GetGalleryImageRowsAsync, GetGalleryImagesHeader, FormatGalleryImageRow);
                return false;
            case "Custom images":
                await ShowGraphRowsAsync("Windows 365 custom images", _session.Graph.GetCustomImageRowsAsync, GetCustomImagesHeader, FormatCustomImageRow);
                return false;
            case "Supported regions":
                await ShowGraphRowsAsync("Windows 365 supported regions", _session.Graph.GetSupportedRegionRowsAsync, GetSupportedRegionsHeader, FormatSupportedRegionRow);
                return false;
        }

        switch (choice.Key)
        {
            case "CommandPalette":
                await ShowCommandPaletteAsync();
                break;
            case "ActionHistory":
                await ShowActionHistoryAsync();
                break;
            case "Connection":
                await ShowConnectionAsync();
                break;
            case "CloudPcs":
                await ShowCloudPcAreaAsync();
                break;
            case "CloudApps":
                await ShowCloudAppsAsync();
                break;
            case "Provisioning":
                await ShowProvisioningAsync();
                break;
            case "Reports":
                await ShowReportsAsync();
                break;
            case "Licensing":
                await ShowLicensingAsync();
                break;
            case "Catalog":
                await ShowCatalogAsync();
                break;
            case "Tenant":
                await ShowTenantSettingsAsync();
                break;
            case "About":
                ShowAbout();
                break;
            case "Exit":
                AnsiConsole.Clear();
                Console.CursorVisible = true;
                return true;
        }

        return false;
    }

    private async Task ShowCloudPcAreaAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var choices = new[] { "Browse Cloud PCs", "Disk space", "Snapshots", "Back" };
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Cloud PCs");
            AnsiConsole.MarkupLine("[#58a6ff]Cloud PCs[/]");
            AnsiConsole.WriteLine();
            for (var index = 0; index < choices.Length; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on #58a6ff]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter select | Esc/B/Q back[/]");
            RenderStatusBar();
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(choices.Length - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = choices.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    switch (choices[selectedIndex])
                    {
                        case "Browse Cloud PCs":
                            await ShowCloudPcsAsync();
                            break;
                        case "Disk space":
                            await ShowDiskSpaceAsync();
                            break;
                        case "Snapshots":
                            await ShowAllSnapshotsAsync();
                            break;
                        case "Back":
                            return;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private async Task ShowDiskSpaceAsync(CloudPcSummary? cloudPc = null)
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var items = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud PC disk space...", async _ =>
            {
                IReadOnlyList<CloudPcSummary>? targets = cloudPc is null ? null : new[] { cloudPc };
                return await _session.Graph.GetCloudPcDiskSpacesAsync(targets);
            });

        if (items.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No disk space rows returned.[/]");
            Pause();
            return;
        }

        if (cloudPc is not null)
        {
            ShowDiskSpaceDetails(items[0]);
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        while (true)
        {
            var visibleItems = FilterDiskSpaces(items, filter);
            if (selectedIndex > visibleItems.Count)
            {
                selectedIndex = visibleItems.Count;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Cloud PCs", "Disk space");
            RenderDiskSpaceTable(items, visibleItems, selectedIndex, filter);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(visibleItems.Count, selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(visibleItems.Count, selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = visibleItems.Count;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Enter:
                    if (selectedIndex == visibleItems.Count)
                    {
                        return;
                    }

                    if (visibleItems.Count > 0)
                    {
                        await OpenCloudPcFromDiskSpaceAsync(visibleItems[selectedIndex]);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private static void RenderDiskSpaceTable(IReadOnlyList<CloudPcDiskSpace> allItems, IReadOnlyList<CloudPcDiskSpace> visibleItems, int selectedIndex, string filter)
    {
        AnsiConsole.MarkupLine("[#58a6ff]Windows 365 Cloud PC disk space[/]");
        AnsiConsole.MarkupLine($"[grey]Rows: {allItems.Count} | Visible: {visibleItems.Count} | Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}[/]");
        var header = Row("Cloud PC", 34, "Free", 10, "Used", 10, "Total", 10, "Free %", 8, "Last sync", 20);
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");
        AnsiConsole.WriteLine();

        var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
        var totalRows = visibleItems.Count + 1;
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, totalRows - pageSize));
        var end = Math.Min(totalRows - 1, start + pageSize - 1);

        for (var index = start; index <= end; index++)
        {
            string label;
            if (index == visibleItems.Count)
            {
                label = "Back";
            }
            else
            {
                var disk = visibleItems[index];
                var hasError = !string.IsNullOrWhiteSpace(disk.Error);
                label = Row(
                    disk.CloudPcName, 34,
                    hasError ? "unavail" : FormatGb(disk.FreeStorageGb), 10,
                    FormatGb(disk.UsedStorageGb), 10,
                    FormatGb(disk.TotalStorageGb), 10,
                    disk.PercentFree is null ? "-" : $"{disk.PercentFree}%", 8,
                    disk.LastSyncDateTime?.ToLocalTime().ToString("g") ?? "-", 20);
            }

            var escaped = Markup.Escape(label);
            AnsiConsole.MarkupLine(index == selectedIndex
                ? $"[black on #58a6ff]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter Cloud PC actions | / or F filter | C clear | Esc/B/Q back[/]");
        RenderStatusBar();
    }

    private static IReadOnlyList<CloudPcDiskSpace> FilterDiskSpaces(IReadOnlyList<CloudPcDiskSpace> items, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return items;
        }

        return items
            .Where(item =>
                Contains(item.CloudPcName, filter) ||
                Contains(item.AssignedUserUpn, filter) ||
                Contains(item.ManagedDeviceName, filter) ||
                Contains(item.ManagedDeviceId, filter) ||
                Contains(item.Error, filter))
            .ToArray();
    }

    private async Task OpenCloudPcFromDiskSpaceAsync(CloudPcDiskSpace disk)
    {
        var cloudPcs = await LoadCloudPcsAsync();
        var cloudPc = cloudPcs.FirstOrDefault(pc =>
            string.Equals(pc.Id, disk.CloudPcId, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.ManagedDeviceId, disk.ManagedDeviceId, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.Name, disk.CloudPcName, StringComparison.OrdinalIgnoreCase));

        if (cloudPc is null)
        {
            ShowDiskSpaceDetails(disk);
            return;
        }

        await ShowCloudPcDetailsAsync(cloudPc);
    }

    private async Task ShowAllSnapshotsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var items = await LoadAllSnapshotsAsync();
        if (items.Count == 0)
        {
            TimedMessage("[yellow]No snapshots were returned.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        while (true)
        {
            var visibleItems = FilterSnapshotItems(items, filter);
            if (visibleItems.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleItems.Count)
            {
                selectedIndex = visibleItems.Count - 1;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Cloud PCs", "Snapshots");
            RenderAllSnapshotsTable(items, visibleItems, selectedIndex, filter);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleItems.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visibleItems.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visibleItems.Count - 1);
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Enter:
                    if (visibleItems.Count > 0)
                    {
                        await ShowCloudPcDetailsAsync(visibleItems[selectedIndex].CloudPc);
                    }
                    break;
                case ConsoleKey.A:
                    if (visibleItems.Count == 0)
                    {
                        break;
                    }

                    await ShowSnapshotActionMenuAsync(visibleItems[selectedIndex].CloudPc, visibleItems[selectedIndex].Snapshot);
                    items = await LoadAllSnapshotsAsync();
                    selectedIndex = Math.Min(selectedIndex, Math.Max(0, items.Count - 1));
                    if (items.Count == 0)
                    {
                        TimedMessage("[yellow]No snapshots were returned.[/]");
                        return;
                    }
                    break;
                case ConsoleKey.R:
                    items = await LoadAllSnapshotsAsync();
                    selectedIndex = 0;
                    if (items.Count == 0)
                    {
                        TimedMessage("[yellow]No snapshots were returned.[/]");
                        return;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private async Task<IReadOnlyList<SnapshotListItem>> LoadAllSnapshotsAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading snapshots...", async _ =>
            {
                var cloudPcs = await _session.Graph.GetCloudPcsAsync();
                var results = new List<SnapshotListItem>();
                foreach (var cloudPc in cloudPcs)
                {
                    var snapshots = await _session.Graph.GetCloudPcSnapshotsAsync(cloudPc);
                    results.AddRange(snapshots.Select(snapshot => new SnapshotListItem(cloudPc, snapshot)));
                }

                return results
                    .OrderByDescending(item => item.Snapshot.CreatedDateTime)
                    .ThenBy(item => item.CloudPc.Name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            });
    }

    private static void RenderAllSnapshotsTable(IReadOnlyList<SnapshotListItem> allItems, IReadOnlyList<SnapshotListItem> visibleItems, int selectedIndex, string filter)
    {
        AnsiConsole.MarkupLine("[#58a6ff]Windows 365 Cloud PC snapshots[/]");
        AnsiConsole.MarkupLine($"[grey]Rows: {allItems.Count} | Visible: {visibleItems.Count} | Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}[/]");
        AnsiConsole.WriteLine();

        var widths = GetAllSnapshotWidths();
        var header = Row("Cloud PC", widths.CloudPc, "Status", widths.Status, "Type", widths.Type, "Created", widths.Created, "Expires", widths.Expires);
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

        if (visibleItems.Count == 0)
        {
            AnsiConsole.MarkupLine("[grey]No snapshots match the current filter.[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]/ or F filter | C clear | Esc/B/Q back[/]");
            return;
        }

        var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visibleItems.Count - pageSize));
        var visible = visibleItems.Skip(start).Take(pageSize).ToArray();
        for (var index = 0; index < visible.Length; index++)
        {
            var item = visible[index];
            var absoluteIndex = start + index;
            var row = Row(
                item.CloudPc.Name, widths.CloudPc,
                item.Snapshot.Status ?? "-", widths.Status,
                item.Snapshot.SnapshotType ?? "-", widths.Type,
                item.Snapshot.CreatedDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Created,
                item.Snapshot.ExpirationDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Expires);
            var escaped = Markup.Escape(row);
            AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                ? $"[black on #58a6ff]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter Cloud PC actions | A snapshot actions | / or F filter | C clear | R refresh | Esc/B/Q back[/]");
        RenderStatusBar();
    }

    private static IReadOnlyList<SnapshotListItem> FilterSnapshotItems(IReadOnlyList<SnapshotListItem> items, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return items;
        }

        return items
            .Where(item =>
                Contains(item.CloudPc.Name, filter) ||
                Contains(item.CloudPc.UserPrincipalName, filter) ||
                Contains(item.Snapshot.Status, filter) ||
                Contains(item.Snapshot.SnapshotType, filter) ||
                Contains(item.Snapshot.HealthCheckStatus, filter) ||
                Contains(item.Snapshot.SnapshotId, filter))
            .ToArray();
    }

    private static (int CloudPc, int Status, int Type, int Created, int Expires) GetAllSnapshotWidths()
    {
        var available = Math.Max(90, Console.WindowWidth - 4);
        const int status = 14;
        const int type = 14;
        const int created = 18;
        const int expires = 18;
        var cloudPc = Math.Max(28, available - status - type - created - expires - 4);
        return (cloudPc, status, type, created, expires);
    }

    private static void ShowDiskSpaceDetails(CloudPcDiskSpace disk)
    {
        AnsiConsole.Clear();
        var panel = new Panel(
            new Rows(
                new Markup($"[bold]Cloud PC:[/] {Markup.Escape(disk.CloudPcName)}"),
                new Markup($"[bold]Managed device:[/] {Markup.Escape(disk.ManagedDeviceName ?? "-")}"),
                new Markup($"[bold]User:[/] {Markup.Escape(disk.AssignedUserUpn ?? "-")}"),
                new Markup($"[bold]Free:[/] {Markup.Escape(FormatGb(disk.FreeStorageGb))}"),
                new Markup($"[bold]Used:[/] {Markup.Escape(FormatGb(disk.UsedStorageGb))}"),
                new Markup($"[bold]Total:[/] {Markup.Escape(FormatGb(disk.TotalStorageGb))}"),
                new Markup($"[bold]Percent free:[/] {Markup.Escape(disk.PercentFree is null ? "-" : $"{disk.PercentFree}%")}"),
                new Markup($"[bold]Last sync:[/] {Markup.Escape(disk.LastSyncDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
                new Markup($"[bold]Status:[/] {Markup.Escape(disk.Error ?? "Disk data available")}"),
                new Markup($"[bold]Cloud PC ID:[/] [grey]{Markup.Escape(disk.CloudPcId)}[/]"),
                new Markup($"[bold]Managed device ID:[/] [grey]{Markup.Escape(disk.ManagedDeviceId ?? "-")}[/]")))
            .Header("Disk space details")
            .Border(BoxBorder.Rounded);
        AnsiConsole.Write(panel);
        WaitForBack();
    }

    private static string FormatGb(double? value)
    {
        return value is null ? "-" : $"{value:0.##} GB";
    }

    private IReadOnlyList<MenuChoice> GetMainMenuChoices()
    {
        var connectionDescription = _session.IsConnected
            ? "Disconnect Microsoft Graph session"
            : "Connect to Microsoft Graph";

        return
        [
            new("CloudPcs", "Cloud PCs", "Browse, inspect, filter, and act on Cloud PCs",
            [
                new("CloudPcs", "Browse Cloud PCs", "Open Cloud PC browser"),
                new("CloudPcs", "Disk space", "Open all Cloud PC disk space"),
                new("CloudPcs", "Snapshots", "Open all Cloud PC snapshots")
            ]),
            new("Provisioning", "Provisioning", "Provisioning policies and maintenance windows",
            [
                new("Provisioning", "Policies", "View, copy, export, reprovision, and delete policies"),
                new("Provisioning", "Create policy", "Create a new provisioning policy"),
                new("Provisioning", "User experience sync overview", "Storage usage across all shared policies with user experience sync")
            ]),
            new("Reports", "Reports", "Usage, connectivity, launch details, report streams",
            [
                new("Reports", "Sign-in status", "Open current Cloud PC sign-in status"),
                new("Reports", "Connectivity history", "Select a Cloud PC and inspect connection events"),
                new("Reports", "Launch details", "Open Cloud PC launch details"),
                new("Reports", "Cloud PC reports", "Browse Graph report streams")
            ]),
            new("Licensing", "Licensing", "Capacity, availability, Flex, and Reserve utilization"),
            new("CloudApps", "Cloud Apps", "Browse, publish, and unpublish Cloud Apps"),
            new("Catalog", "Catalog", "Service plans, images, regions",
            [
                new("Catalog", "Service plans", "Open service plan catalog"),
                new("Catalog", "Gallery images", "Open gallery images"),
                new("Catalog", "Custom images", "Open custom images"),
                new("Catalog", "Supported regions", "Open supported regions")
            ]),
            new("Tenant", "Tenant settings", "Organization settings, profiles, user settings",
            [
                new("Tenant", "Organization settings", "View tenant-wide Windows 365 defaults"),
                new("Tenant", "Setting profiles", "View Windows 365 setting profiles"),
                new("Tenant", "User settings", "View user settings policies")
            ]),
            new("Connection", "Connection", connectionDescription),
            new("About", "About", "Version and project information"),
            new("Exit", "Exit", "Close W365 CLI")
        ];
    }

    private void RenderMainMenuDashboard(IReadOnlyList<MenuChoice> choices)
    {
        UpdateStatusBarSnapshot();
    }

    private void RenderMainMenu(IReadOnlyList<MenuChoice> choices, int selectedIndex, int expandedIndex, int selectedChildIndex, int topNavIndex)
    {
        var compact = IsCompactLayout();
        RenderHeader(focusedTopNavIndex: topNavIndex);
        if (!compact)
        {
            RenderTip();
            AnsiConsole.WriteLine();
        }

        RenderHomeStatusLine();
        if (!compact)
        {
            AnsiConsole.WriteLine();
        }

        RenderMainMenuDashboard(choices);
        var isConnected = _session.IsConnected;
        var menuRows = new List<Markup>();
        for (var index = 0; index < choices.Count; index++)
        {
            var selected = index == selectedIndex;
            var disabled = IsMenuChoiceDisabledWhenDisconnected(choices[index], isConnected);
            var label = FormatMainMenuChoice(choices[index], selected, disabled);
            var expandMarker = choices[index].Children?.Count > 0
                ? expandedIndex == index ? "v " : "  "
                : "  ";
            menuRows.Add(new Markup(selected
                ? disabled
                    ? $"[#484f58 on #21262d]{expandMarker}{label}[/]"
                    : $"[black on {AccentColor}]{expandMarker}{label}[/]"
                : $"{expandMarker}{label}"));

            if (expandedIndex == index && choices[index].Children is { Count: > 0 } children)
            {
                for (var childIndex = 0; childIndex < children.Count; childIndex++)
                {
                    var childSelected = selected && selectedChildIndex == childIndex;
                    var childDisabled = IsMenuChoiceDisabledWhenDisconnected(children[childIndex], isConnected);
                    var childLabel = FormatMainMenuChoice(children[childIndex], childSelected, childDisabled);
                    menuRows.Add(new Markup(childSelected
                        ? childDisabled
                            ? $"[#484f58 on #21262d]  > {childLabel}[/]"
                            : $"[black on {AccentColor}]  > {childLabel}[/]"
                        : $"    {childLabel}"));
                }
            }
        }

        AnsiConsole.Write(new Panel(new Rows(menuRows))
            .Header("Main menu")
            .Border(BoxBorder.Rounded)
            .BorderStyle(new Style(Color.FromHex(AccentColor)))
            .Expand());
        AnsiConsole.WriteLine();
        var hint = isConnected
            ? "[grey]Tab top nav | Up/Down move | Right expand | Enter select | Esc/B/Q collapse | P or Ctrl+K command palette[/]"
            : "[grey]Not connected — most areas are unavailable until you connect. Select Connection to sign in.[/]";
        RenderTopNavAwareHint(topNavIndex, hint);
    }

    /// <summary>
    /// Every main-menu area except Connection/About/Exit calls through to Microsoft Graph, so
    /// they're all non-functional until the session is connected. Greying them out up front (via
    /// this check) avoids the user drilling into a screen only to be prompted to connect there —
    /// the same connection gate the screens already enforce internally (EnsureConnectedAsync)
    /// still applies as a fallback if they select one anyway.
    /// </summary>
    private static bool IsMenuChoiceDisabledWhenDisconnected(MenuChoice choice, bool isConnected)
    {
        if (isConnected)
        {
            return false;
        }

        return choice.Key is not ("Connection" or "About" or "Exit");
    }

    private void RenderTip()
    {
        var body = $"[bold {PurpleColor}]Tip: {Markup.Escape(launchTip.Command)}[/]\n" +
            $"[{MutedColor}]L {Markup.Escape(launchTip.Text)}[/]";
        AnsiConsole.Write(new Panel(new Markup(body))
            .Border(BoxBorder.Rounded)
            .BorderStyle(new Style(Color.FromHex(PurpleColor)))
            .Padding(0, 0, 0, 0));
    }

    private static void RenderHomeStatusLine()
    {
        var transient = statusMessage is not null &&
            statusMessageAt is not null &&
            DateTimeOffset.Now - statusMessageAt < TimeSpan.FromSeconds(6)
                ? $"  [{MutedColor}]|[/]  {statusMessage}"
                : string.Empty;

        AnsiConsole.MarkupLine($"Graph {statusBarConnection}  [{MutedColor}]|[/]  Tenant [{TextColor}]{Markup.Escape(statusBarTenant)}[/]{transient}");
    }

    private async Task ShowCommandPaletteAsync()
    {
        var commands = GetCommandPaletteChoices();
        var selectedIndex = 0;
        var filter = string.Empty;
        while (true)
        {
            var visibleCommands = FilterMenuChoices(commands, filter);
            if (visibleCommands.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleCommands.Count)
            {
                selectedIndex = visibleCommands.Count - 1;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Command palette");
            AnsiConsole.MarkupLine("[#58a6ff]Command palette[/]");
            AnsiConsole.MarkupLine($"[grey]Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}[/]");
            AnsiConsole.WriteLine();

            if (visibleCommands.Count == 0)
            {
                AnsiConsole.MarkupLine("[grey]No commands match the current filter.[/]");
            }
            else
            {
                var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 12));
                var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visibleCommands.Count - pageSize));
                var visiblePage = visibleCommands.Skip(start).Take(pageSize).ToArray();
                for (var index = 0; index < visiblePage.Length; index++)
                {
                    var absoluteIndex = start + index;
                    var label = FormatMainMenuChoice(visiblePage[index]);
                    AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                        ? $"[black on #58a6ff]> {label}[/]"
                        : $"  {label}");
                }
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Type to filter | Up/Down move | Enter run | Backspace edit | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleCommands.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visibleCommands.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Enter:
                    if (visibleCommands.Count > 0)
                    {
                        await ExecuteMainMenuChoiceAsync(visibleCommands[selectedIndex]);
                        return;
                    }
                    break;
                case ConsoleKey.Backspace:
                    if (filter.Length > 0)
                    {
                        filter = filter[..^1];
                        selectedIndex = 0;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (string.IsNullOrWhiteSpace(filter) && key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }

                    if (!char.IsControl(key.KeyChar))
                    {
                        filter += key.KeyChar;
                        selectedIndex = 0;
                    }
                    break;
            }
        }
    }

    private static string FormatMainMenuChoice(MenuChoice choice, bool selected = false, bool disabled = false)
    {
        if (disabled)
        {
            return $"[{MutedColor}]{Markup.Escape(Fit(choice.Title, 22))}[/] [{MutedColor}]{Markup.Escape(choice.Description)}[/]";
        }

        var descriptionColor = selected ? TextColor : MutedColor;
        return $"[{TextColor}]{Markup.Escape(Fit(choice.Title, 22))}[/] [{descriptionColor}]{Markup.Escape(choice.Description)}[/]";
    }

    private static void RenderBreadcrumb(params string[] parts)
    {
        RenderTopNav();
        var allParts = new[] { "W365 CLI" }
            .Concat(parts.Where(part => !string.IsNullOrWhiteSpace(part)))
            .ToArray();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(string.Join(" > ", allParts))}[/]");
        AnsiConsole.WriteLine();
    }

    private IReadOnlyList<MenuChoice> GetCommandPaletteChoices()
    {
        return
        [
            ..GetMainMenuChoices().Where(choice => choice.Key != "Exit"),
            new("CloudPcs", "Browse Cloud PCs", "Open Cloud PC browser"),
            new("CloudPcs", "Disk space", "Open all Cloud PC disk space"),
            new("CloudPcs", "Snapshots", "Open all Cloud PC snapshots"),
            new("Provisioning", "Policies", "Open provisioning policy browser"),
            new("Provisioning", "Create policy", "Create a new provisioning policy"),
            new("Provisioning", "User experience sync overview", "Storage usage across all shared policies with user experience sync"),
            new("Reports", "Usage report", "Open Cloud PC usage"),
            new("Licensing", "Licensing", "Open licensing capacity view"),
            new("Reports", "Launch details", "Open Cloud PC launch details"),
            new("Catalog", "Service plans", "Open service plan catalog"),
            new("Catalog", "Gallery images", "Open gallery images"),
            new("Catalog", "Supported regions", "Open supported regions")
        ];
    }

    private static IReadOnlyList<MenuChoice> FilterMenuChoices(IReadOnlyList<MenuChoice> choices, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return choices;
        }

        return choices
            .Where(choice => Contains(choice.Title, filter) || Contains(choice.Description, filter))
            .ToArray();
    }

    private static void RenderTopNav(string? active = null, int focusedIndex = -1)
    {
        string Tab(string name, int index)
        {
            var isActive = focusedIndex == index || (focusedIndex < 0 && string.Equals(active, name, StringComparison.OrdinalIgnoreCase));
            return isActive ? $"[{name}]" : $" {name} ";
        }

        var navText = $"  {Tab("Home", 0)}     {Tab("About", 1)}     {Tab("Exit", 2)}";
        var width = Math.Max(navText.Length, Console.WindowWidth - 1);
        AnsiConsole.MarkupLine($"[black on {AccentColor}]{Markup.Escape(navText.PadRight(width))}[/]");
        AnsiConsole.WriteLine();
    }

    /// <summary>
    /// Renders the keyboard hint line for a screen that shows the top nav. When a top-nav tab is
    /// focused (topNavIndex >= 0), this replaces the screen's normal hint with an unambiguous
    /// call to action so the user always knows what pressing Enter will do — most importantly,
    /// making it obvious that focusing "Exit" requires a follow-up Enter to actually quit.
    /// </summary>
    private static void RenderTopNavAwareHint(int topNavIndex, string defaultHint)
    {
        var hint = topNavIndex switch
        {
            0 => "[black on yellow] Home [/] [grey]selected — press[/] [bold]Enter[/] [grey]to return to the main menu, or[/] [bold]Esc[/][grey]/[/][bold]Tab[/] [grey]to keep going[/]",
            1 => "[black on yellow] About [/] [grey]selected — press[/] [bold]Enter[/] [grey]to open About, or[/] [bold]Esc[/][grey]/[/][bold]Tab[/] [grey]to keep going[/]",
            2 => "[black on yellow] Exit [/] [grey]selected — press[/] [bold]Enter[/] [grey]to quit W365 CLI, or[/] [bold]Esc[/][grey]/[/][bold]Tab[/] [grey]to keep going[/]",
            _ => defaultHint
        };

        AnsiConsole.MarkupLine(hint);
    }

    /// <summary>
    /// The full home screen (top nav + banner + tip + status + main menu panel) needs roughly
    /// 34-36 rows to render without clipping. Windows Terminal's default profile height (and many
    /// other terminals' defaults) is shorter than that, which silently scrolls the top of the
    /// frame out of view on every redraw — the user never sees it cut off "once", it's cut off
    /// every single time. Rather than picking one arbitrary width, react to the terminal's actual
    /// reported height and switch to a shorter layout that fits, keyed off <see cref="CompactRowsThreshold"/>.
    /// </summary>
    private const int CompactRowsThreshold = 34;

    private static bool IsCompactLayout()
    {
        try
        {
            return Console.WindowHeight > 0 && Console.WindowHeight < CompactRowsThreshold;
        }
        catch
        {
            // Console.WindowHeight can throw when output is redirected (e.g. piped/non-interactive) —
            // fall back to the full layout in that case since there's no terminal size to react to.
            return false;
        }
    }

    private void RenderHeader(string? activeNav = "Home", int focusedTopNavIndex = -1)
    {
        AnsiConsole.Clear();
        UpdateStatusBarSnapshot();
        RenderTopNav(activeNav, focusedTopNavIndex);

        if (IsCompactLayout())
        {
            AnsiConsole.MarkupLine($"[bold {AccentColor}]W365 CLI[/] [{MutedColor}]v{GetCurrentVersion()} | Bradley Wyatt[/]");
            return;
        }

        AnsiConsole.Write(new Panel(new Rows(
                new Markup($"[{AccentColor}]██╗    ██╗██████╗  ██████╗ ███████╗     ██████╗██╗     ██╗[/]"),
                new Markup($"[{AccentColor}]██║    ██║╚════██╗██╔════╝ ██╔════╝    ██╔════╝██║     ██║[/]"),
                new Markup($"[{AccentColor}]██║ █╗ ██║ █████╔╝███████╗ ███████╗    ██║     ██║     ██║[/]"),
                new Markup($"[{AccentColor}]██║███╗██║ ╚═══██╗██╔═══██╗╚════██║    ██║     ██║     ██║[/]"),
                new Markup($"[{AccentColor}]╚███╔███╔╝██████╔╝╚██████╔╝███████║    ╚██████╗███████╗██║[/]"),
                new Markup($"[{AccentColor}] ╚══╝╚══╝ ╚═════╝  ╚═════╝ ╚══════╝     ╚═════╝╚══════╝╚═╝[/]"),
                new Markup(""),
                new Markup($"[{MutedColor}]Version: v{GetCurrentVersion()} | Author: Bradley Wyatt[/]")))
            .Border(BoxBorder.Rounded)
            .BorderStyle(new Style(Color.FromHex(AccentColor)))
            .Expand());
        AnsiConsole.WriteLine();
    }

    private async Task ShowConnectionAsync()
    {
        RenderHeader(activeNav: null);

        var choices = _session.IsConnected
            ? new[] { "Disconnect", "Back" }
            : new[] { "Connect", "Back" };

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("[#58a6ff]Connection[/]")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices(choices));

        switch (choice)
        {
            case "Connect":
                await _session.ConnectAsync();
                UpdateStatusBarSnapshot();
                await ShowMissingPermissionPromptIfNeededAsync();
                TimedMessage("[grey]Returning...[/]");
                break;
            case "Disconnect":
                await _session.DisconnectAsync();
                UpdateStatusBarSnapshot();
                TimedMessage("[green]Disconnected.[/]");
                break;
        }
    }

    /// <summary>
    /// Surfaces a first-run-friendly prompt right after connecting if any of the app's required
    /// Graph permissions (<see cref="W365Session.MissingRequiredScopes"/>) aren't granted in this
    /// tenant yet, so the user can fix it up front instead of discovering it later as a confusing
    /// 403 mid-action.
    /// </summary>
    private Task ShowMissingPermissionPromptIfNeededAsync()
    {
        if (!_session.IsConnected || _session.MissingRequiredScopes.Count == 0)
        {
            return Task.CompletedTask;
        }

        AnsiConsole.Clear();
        AnsiConsole.MarkupLine("[yellow]Heads up: this app registration is missing permission(s) in your tenant that some features rely on:[/]");
        foreach (var scope in _session.MissingRequiredScopes)
        {
            AnsiConsole.MarkupLine($"[grey]  - {Markup.Escape(scope)}[/]");
        }
        AnsiConsole.MarkupLine("[grey]Features that use these will fail with 403 Forbidden until a Global or Cloud Application administrator adds and grants consent for them.[/]");
        AnsiConsole.WriteLine();

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("How would you like to proceed?")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Open admin consent page now", "Continue anyway"));

        if (choice == "Open admin consent page now")
        {
            OpenUrl(_session.GetAdminConsentUrl());
            TimedMessage("[grey]Opened the admin consent page in your browser. Have an admin approve it, then reconnect.[/]");
        }

        return Task.CompletedTask;
    }

    private static void ShowPlaceholderArea(string title, string message)
    {
        AnsiConsole.Clear();
        var panel = new Panel(
            new Rows(
                new Markup($"[bold]{Markup.Escape(title)}[/]"),
                new Markup(Markup.Escape(message)),
                new Markup("[grey]This area exists in the PowerShell CLI and is queued for native implementation.[/]")))
            .Header(title)
            .Border(BoxBorder.Rounded);

        AnsiConsole.Write(panel);
        Pause();
    }

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

    [System.Runtime.Versioning.SupportedOSPlatform("macos")]
    private static async Task DownloadAndInstallMacUpdateAsync(GitHubReleaseInfo release)
    {
        var asset = FindMacZipAsset(release);
        if (asset is null)
        {
            AnsiConsole.MarkupLine($"[yellow]Couldn't find a macOS build for this release ({Markup.Escape(GetCurrentOsArch())}). Opening the release page instead.[/]");
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
            var zipPath = Path.Combine(tempDir, asset.Name);

            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Downloading {asset.Name}...", async _ => await DownloadFileAsync(asset.BrowserDownloadUrl, zipPath));

            var extractDir = Path.Combine(tempDir, "extracted");
            Directory.CreateDirectory(extractDir);
            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractDir);

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
            // terminal's tty/termios settings in a bad state on macOS — since Console.ReadKey's
            // raw-mode handling doesn't get a chance to clean up on an abrupt exit — which then
            // makes the *next* process launched in that same terminal window crash with an
            // Input/output error the moment it tries to read a key. The binary on disk is already
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

                await DownloadAndInstallMacUpdateAsync(release);
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

    private async Task ShowProvisioningAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var policies = await LoadProvisioningPoliciesAsync();
        if (policies.Count == 0)
        {
            TimedMessage("[yellow]No provisioning policies were returned. Use \"Create policy\" to add one.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = ProvisioningPolicySortMode.Name;
        while (true)
        {
            var visiblePolicies = SortProvisioningPolicies(FilterProvisioningPolicies(policies, filter), sortMode);
            if (visiblePolicies.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visiblePolicies.Count)
            {
                selectedIndex = visiblePolicies.Count - 1;
            }

            AnsiConsole.Clear();
            RenderProvisioningPolicyBrowser(policies, visiblePolicies, selectedIndex, filter, sortMode);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visiblePolicies.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visiblePolicies.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visiblePolicies.Count - 1);
                    break;
                case ConsoleKey.Enter:
                case ConsoleKey.A:
                    if (visiblePolicies.Count > 0)
                    {
                        await ShowProvisioningPolicyDetailsAsync(visiblePolicies[selectedIndex]);
                        policies = await LoadProvisioningPoliciesAsync();
                    }
                    break;
                case ConsoleKey.R:
                    policies = await LoadProvisioningPoliciesAsync();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextProvisioningPolicySortMode(sortMode);
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

        private async Task ShowLicensingAsync()
        {
            if (!await EnsureConnectedAsync())
            {
                return;
            }

            IReadOnlyList<LicenseOverviewItem> items;
            try
            {
                items = await LoadLicenseOverviewAsync();
            }
            catch (Exception ex)
            {
                AnsiConsole.Clear();
                RenderBreadcrumb("Licensing");
                AnsiConsole.MarkupLine("[red]Failed to load licensing data.[/]");
                AnsiConsole.MarkupLine("[grey]The Licensing view requires access to subscribedSkUs, usually via Organization.Read.All or equivalent directory licensing permissions.[/]");
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[grey]{Markup.Escape(ex.Message)}[/]");
                WaitForBack();
                return;
            }

            if (items.Count == 0)
            {
                TimedMessage("[yellow]No Windows 365 license SKUs were detected from subscribedSkUs.[/]");
                return;
            }

            var selectedIndex = 0;
            while (true)
            {
                AnsiConsole.Clear();
                RenderLicensingOverview(items, selectedIndex);
                var key = ReadNavigationKey(intercept: true);
                switch (key.Key)
                {
                    case ConsoleKey.UpArrow:
                        selectedIndex = Math.Max(0, selectedIndex - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        selectedIndex = Math.Min(items.Count - 1, selectedIndex + 1);
                        break;
                    case ConsoleKey.PageUp:
                        selectedIndex = Math.Max(0, selectedIndex - 10);
                        break;
                    case ConsoleKey.PageDown:
                        selectedIndex = Math.Min(items.Count - 1, selectedIndex + 10);
                        break;
                    case ConsoleKey.Home:
                        selectedIndex = 0;
                        break;
                    case ConsoleKey.End:
                        selectedIndex = items.Count - 1;
                        break;
                    case ConsoleKey.Enter:
                        await ShowLicenseDetailsAsync(items[selectedIndex]);
                        break;
                    case ConsoleKey.R:
                        items = await LoadLicenseOverviewAsync();
                        selectedIndex = Math.Min(selectedIndex, Math.Max(0, items.Count - 1));
                        if (items.Count == 0)
                        {
                            TimedMessage("[yellow]No Windows 365 license SKUs were detected from subscribedSkUs.[/]");
                            return;
                        }
                        break;
                    case ConsoleKey.Escape:
                    case ConsoleKey.LeftArrow:
                        return;
                    default:
                        if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                        {
                            return;
                        }
                        break;
                }
            }
        }

        private async Task<IReadOnlyList<LicenseOverviewItem>> LoadLicenseOverviewAsync()
        {
            return await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading license data...", async _ =>
                {
                    var skus = await _session.Graph.GetSubscribedSkusAsync();
                    var cloudPcs = await _session.Graph.GetCloudPcsAsync();
                    var policies = await _session.Graph.GetProvisioningPoliciesAsync();
                    return BuildLicenseOverview(skus, cloudPcs, policies);
                });
        }

        private static IReadOnlyList<LicenseOverviewItem> BuildLicenseOverview(
            IReadOnlyList<SubscribedSku> skus,
            IReadOnlyList<CloudPcSummary> cloudPcs,
            IReadOnlyList<ProvisioningPolicySummary> policies)
        {
            var windows365Skus = skus
                .Select(sku => new { Sku = sku, Info = GetWindows365LicenseInfo(sku) })
                .Where(item => item.Info is not null)
                .ToArray();
            var output = new List<LicenseOverviewItem>();
            foreach (var group in windows365Skus
                .GroupBy(item => $"{item.Info!.Family}|{item.Info.PlanKey}", StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.First().Info!.Family, StringComparer.OrdinalIgnoreCase)
                .ThenBy(group => group.First().Info!.PlanKey, StringComparer.OrdinalIgnoreCase))
            {
                var info = group.First().Info!;
                var family = info.Family;
                var purchased = group.Sum(item => item.Sku.PrepaidUnits?.Enabled ?? 0);
                var assigned = group.Sum(item => item.Sku.ConsumedUnits ?? 0);
                var isFlex = family.Equals("Flex", StringComparison.OrdinalIgnoreCase);
                var matchingCloudPcs = GetCloudPcsForLicense(cloudPcs, info);
                var dedicated = isFlex ? matchingCloudPcs.Count(pc => GetFlexCloudPcMode(pc) == "Dedicated") : 0;
                var shared = isFlex ? matchingCloudPcs.Count - dedicated : 0;
                var dedicatedUnitsUsed = isFlex ? (int)Math.Ceiling(dedicated / 3d) : 0;
                var sharedUnitsUsed = isFlex ? shared : 0;
                var licenseUnitsUsed = isFlex ? sharedUnitsUsed + dedicatedUnitsUsed : matchingCloudPcs.Count;
                var licenseUnitsLeft = Math.Max(0, purchased - licenseUnitsUsed);
                var provisionable = isFlex ? purchased * 3 : purchased;
                var activeLimit = isFlex ? purchased : purchased;
                var available = isFlex
                    ? licenseUnitsLeft * 3 + Math.Max(0, dedicatedUnitsUsed * 3 - dedicated)
                    : Math.Max(0, provisionable - matchingCloudPcs.Count);
                var flexPolicies = policies
                    .Where(policy => isFlex && IsFlexPolicy(policy))
                    .ToArray();

                output.Add(new LicenseOverviewItem(
                    info.DisplayName,
                    string.Join(", ", group.Select(item => item.Sku.SkuPartNumber).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase)),
                    purchased,
                    assigned,
                    matchingCloudPcs.Count,
                    dedicated,
                    shared,
                    provisionable,
                    available,
                    activeLimit,
                    dedicatedUnitsUsed,
                    sharedUnitsUsed,
                    licenseUnitsUsed,
                    licenseUnitsLeft,
                    matchingCloudPcs,
                    flexPolicies));
            }

            return output;
        }

        private static Windows365LicenseInfo? GetWindows365LicenseInfo(SubscribedSku sku)
        {
            var text = $"{sku.SkuPartNumber} {string.Join(' ', sku.ServicePlans?.Select(plan => plan.ServicePlanName) ?? [])}";
            if (!text.Contains("W365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("WINDOWS_365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("WINDOWS365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CPC_", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CLOUD_PC", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CLOUDPC", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (text.Contains("DISASTER", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("ADD_ON", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("ADDON", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var family = IsReserveText(text)
                ? "Reserve"
                : text.Contains("FLEX", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("FRONTLINE", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("CPC_F_", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("CPC_S_", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("WINDOWS_365_S_", StringComparison.OrdinalIgnoreCase)
                    ? "Flex"
                    : text.Contains("BUSINESS", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("CPC_B_", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("WINDOWS_365_B_", StringComparison.OrdinalIgnoreCase)
                        ? "Business"
                        : "Enterprise";
            var planKey = GetPlanKey(text) ?? "unknown";
            var planLabel = planKey == "unknown" ? "unknown size" : FormatPlanKey(planKey);
            return new Windows365LicenseInfo(family, planKey, $"{family} {planLabel}");
        }

        private static IReadOnlyList<CloudPcSummary> GetCloudPcsForLicense(IReadOnlyList<CloudPcSummary> cloudPcs, Windows365LicenseInfo license)
        {
            return cloudPcs
                .Where(pc =>
                    (license.PlanKey.Equals("unknown", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(GetPlanKey(pc.ServicePlanName ?? string.Empty), license.PlanKey, StringComparison.OrdinalIgnoreCase)) &&
                    IsCloudPcInLicenseFamily(pc, license.Family))
                .ToArray();
        }

        private static bool IsCloudPcInLicenseFamily(CloudPcSummary pc, string family)
        {
            if (family.Equals("Flex", StringComparison.OrdinalIgnoreCase))
            {
                return Contains(pc.ServicePlanName, "Frontline") ||
                    Contains(pc.ServicePlanName, "Flex") ||
                    Contains(pc.ProvisioningPolicyName, "Flex");
            }

            if (family.Equals("Reserve", StringComparison.OrdinalIgnoreCase))
            {
                return IsReserveText($"{pc.ServicePlanName} {pc.ProvisioningType} {pc.ProvisioningPolicyName}");
            }

            return Contains(pc.ServicePlanName, family);
        }

        private static bool IsReserveText(string? value)
        {
            return Contains(value, "Reserve") ||
                Contains(value, "CPC_R_") ||
                Contains(value, "WINDOWS_365_R_");
        }

        private static string? GetPlanKey(string value)
        {
            var displayMatch = System.Text.RegularExpressions.Regex.Match(
                value,
                @"(?<cpu>\d+)\s*vCPU[^\d]+(?<ram>\d+)\s*GB[^\d]+(?<storage>\d+)\s*GB",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (displayMatch.Success)
            {
                return $"{displayMatch.Groups["cpu"].Value}/{displayMatch.Groups["ram"].Value}/{displayMatch.Groups["storage"].Value}";
            }

            var skuMatch = System.Text.RegularExpressions.Regex.Match(
                value,
                @"(?<cpu>\d+)\s*C[^\d]+(?<ram>\d+)\s*GB[^\d]+(?<storage>\d+)\s*GB",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return skuMatch.Success
                ? $"{skuMatch.Groups["cpu"].Value}/{skuMatch.Groups["ram"].Value}/{skuMatch.Groups["storage"].Value}"
                : null;
        }

        private static string FormatPlanKey(string planKey)
        {
            var parts = planKey.Split('/');
            return parts.Length == 3 ? $"{parts[0]}vCPU/{parts[1]}GB/{parts[2]}GB" : planKey;
        }

        private static bool IsDedicatedCloudPc(CloudPcSummary cloudPc)
        {
            return cloudPc.ProvisioningType?.Contains("dedicated", StringComparison.OrdinalIgnoreCase) == true ||
                cloudPc.ProvisioningPolicyName?.Contains("dedicated", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static bool IsSharedCloudPc(CloudPcSummary cloudPc)
        {
            return cloudPc.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true ||
                cloudPc.ProvisioningPolicyName?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static bool IsFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true ||
                policy.DisplayName.Contains("Flex", StringComparison.OrdinalIgnoreCase) ||
                policy.Description?.Contains("Flex", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static void RenderLicensingOverview(IReadOnlyList<LicenseOverviewItem> items, int selectedIndex)
        {
            RenderBreadcrumb("Licensing");
            AnsiConsole.MarkupLine("[#58a6ff]Windows 365 licensing[/]");
            AnsiConsole.MarkupLine("[grey]Capacity estimates use Microsoft Graph subscribedSkUs plus current Cloud PC inventory.[/]");
            AnsiConsole.WriteLine();
            var table = new Table()
                .Border(TableBorder.Rounded)
                .AddColumn(" ")
                .AddColumn("Family")
                .AddColumn("Purchased")
                .AddColumn("Assigned")
                .AddColumn("Cloud PCs")
                .AddColumn("Dedicated")
                .AddColumn("Shared")
                .AddColumn("Units used")
                .AddColumn("Units left")
                .AddColumn("Can run now");

            for (var index = 0; index < items.Count; index++)
            {
                var item = items[index];
                table.AddRow(
                    index == selectedIndex ? "[black on #58a6ff]>[/]" : " ",
                    Markup.Escape(item.Family),
                    item.Purchased.ToString(),
                    item.Assigned.ToString(),
                    item.CloudPcCount.ToString(),
                    item.DedicatedCloudPcCount.ToString(),
                    item.SharedCloudPcCount.ToString(),
                    item.LicenseUnitsUsed.ToString(),
                    item.LicenseUnitsLeft.ToString(),
                    item.ActiveSessionLimit.ToString());
            }

            AnsiConsole.Write(table);
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Enter details | R refresh | Esc/B/Q back[/]");
        }

        private async Task ShowLicenseDetailsAsync(LicenseOverviewItem item)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Licensing", item.Family);
            var detailRows = new List<Markup>
            {
                new(PropertyInline("Family", item.Family)),
                new(PropertyBlock("SKUs", item.SkuPartNumbers)),
                new(PropertyInline("Purchased licenses", item.Purchased.ToString())),
                new(PropertyInline("Assigned licenses", item.Assigned.ToString()))
            };
            if (IsFlexLicense(item))
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Dedicated machines you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Dedicated machines already created", item.DedicatedCloudPcCount.ToString())),
                    new Markup(PropertyInline("More dedicated machines you can create", item.AvailableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Shared pool Cloud PCs", item.SharedCloudPcCount.ToString())),
                    new Markup(PropertyInline("Total Flex Cloud PCs visible", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("Shared license units used", item.SharedUnitsUsed.ToString())),
                    new Markup(PropertyInline("Dedicated license units used", item.DedicatedUnitsUsed.ToString())),
                    new Markup(PropertyInline("Total license units used", item.LicenseUnitsUsed.ToString())),
                    new Markup(PropertyInline("License units left", item.LicenseUnitsLeft.ToString())),
                    new Markup(PropertyInline("Flex Cloud PCs that can have a user connected at once", item.ActiveSessionLimit.ToString()))
                ]);
            }
            else if (IsReserveLicense(item))
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Reserve Cloud PCs you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Reserve Cloud PCs already created", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("More Reserve Cloud PCs you can create", item.AvailableCloudPcCount.ToString()))
                ]);
            }
            else
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Cloud PCs you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Cloud PCs already created", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("More Cloud PCs you can create", item.AvailableCloudPcCount.ToString()))
                ]);
            }
            var details = new Rows(detailRows);
            AnsiConsole.Write(new Panel(details).Header("License capacity").Border(BoxBorder.Rounded));
            AnsiConsole.WriteLine();
            AnsiConsole.Write(CreateLicensePlainEnglishPanel(item));

            if (IsFlexLicense(item))
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[#58a6ff]Flex rules[/]");
                AnsiConsole.MarkupLine($"[grey]Each Flex license unit can cover either 1 shared pool Cloud PC or up to 3 dedicated Cloud PCs. You have {item.LicenseUnitsLeft} license units left.[/]");
                AnsiConsole.WriteLine();
                var groupMembers = await LoadFlexPolicyGroupMembersAsync(item);
                AnsiConsole.Write(BuildCloudAppsPoolTable(item, groupMembers));
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildSharedPoolTable(item, groupMembers));
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildDedicatedMachineTable(item, groupMembers));
            }
            else
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildLicenseCloudPcTable(item));
            }

            WaitForBack();
        }

        private static bool IsFlexLicense(LicenseOverviewItem item)
        {
            return item.Family.StartsWith("Flex", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsReserveLicense(LicenseOverviewItem item)
        {
            return item.Family.StartsWith("Reserve", StringComparison.OrdinalIgnoreCase);
        }

        private static Panel CreateLicensePlainEnglishPanel(LicenseOverviewItem item)
        {
            string[] lines;
            if (IsFlexLicense(item))
            {
                lines = new[]
                {
                    $"You bought {item.Purchased} Flex licenses for this size.",
                    $"{item.SharedCloudPcCount} shared pool or Cloud Apps Cloud PCs use {item.SharedUnitsUsed} license units.",
                    $"{item.DedicatedCloudPcCount} dedicated Cloud PCs use {item.DedicatedUnitsUsed} license units because dedicated capacity is counted in groups of 3.",
                    $"That uses {item.LicenseUnitsUsed} of {item.Purchased} license units, leaving {item.LicenseUnitsLeft}.",
                    $"You can create {item.AvailableCloudPcCount} more dedicated Cloud PCs. You can create {item.LicenseUnitsLeft} more shared pool Cloud PCs.",
                    $"Concurrency means {item.ActiveSessionLimit} total Flex Cloud PCs can have a user connected at the same time. It does not mean multiple users can connect to one Cloud PC."
                };
            }
            else if (IsReserveLicense(item))
            {
                lines =
                [
                    $"You bought {item.Purchased} Reserve licenses for this size.",
                    $"{item.CloudPcCount} Reserve Cloud PCs currently match this license size.",
                    $"{item.AvailableCloudPcCount} more Reserve Cloud PCs can be created before this license capacity is used.",
                    "The Cloud PC table below shows which user is assigned to each detected Reserve Cloud PC."
                ];
            }
            else
            {
                lines = new[]
                {
                    $"You bought {item.Purchased} licenses for this size.",
                    $"{item.CloudPcCount} Cloud PCs currently match this license size.",
                    $"{item.AvailableCloudPcCount} more Cloud PCs can be created before this license capacity is used."
                };
            }

            return new Panel(new Rows(lines.Select(line => new Markup($"[grey]{Markup.Escape(line)}[/]"))))
                .Header("What this means")
                .Border(BoxBorder.Rounded);
        }

        private static Table BuildLicenseCloudPcTable(LicenseOverviewItem item)
        {
            var table = new Table()
                .Border(TableBorder.Rounded)
                .AddColumn("Cloud PC")
                .AddColumn("Mode")
                .AddColumn("Assigned user")
                .AddColumn("Service plan")
                .AddColumn("Policy");

            foreach (var cloudPc in item.CloudPcs.OrderBy(pc => pc.ProvisioningPolicyName).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase))
            {
                table.AddRow(
                    Markup.Escape(cloudPc.Name),
                    Markup.Escape(GetFlexCloudPcMode(cloudPc)),
                    Markup.Escape(cloudPc.UserPrincipalName ?? "-"),
                    Markup.Escape(cloudPc.ServicePlanName ?? "-"),
                    Markup.Escape(cloudPc.ProvisioningPolicyName ?? "-"));
            }

            if (item.CloudPcs.Count == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No matching Cloud PCs detected.[/]");
            }

            return table;
        }

        private async Task<IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>>> LoadFlexPolicyGroupMembersAsync(LicenseOverviewItem item)
        {
            var output = new Dictionary<string, IReadOnlyList<GroupMemberSummary>>(StringComparer.OrdinalIgnoreCase);
            foreach (var groupId in item.FlexPolicies.SelectMany(policy => policy.AssignedGroupIds).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    output[groupId] = await _session.Graph.GetGroupMembersAsync(groupId);
                }
                catch (Exception)
                {
                    output[groupId] = [];
                }
            }

            return output;
        }

        private static Table BuildCloudAppsPoolTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Cloud Apps pool")
                .Border(TableBorder.Rounded)
                .AddColumn("Policy")
                .AddColumn("Group")
                .AddColumn("Cloud PC")
                .AddColumn("User")
                .AddColumn("UPN");

            var policies = item.FlexPolicies.Where(IsCloudAppsPolicy).ToArray();
            foreach (var policy in policies)
            {
                var cloudPcs = item.CloudPcs.Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase)).ToArray();
                var cloudPcList = cloudPcs.Length == 0 ? "-" : string.Join(", ", cloudPcs.Select(pc => pc.Name));
                foreach (var accessRow in GetPolicyAccessRows(policy, groupMembers))
                {
                    table.AddRow(
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(accessRow.GroupName),
                        Markup.Escape(cloudPcList),
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName));
                }
            }

            if (policies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No Cloud Apps pool detected.[/]", "[grey]-[/]");
            }

            return table;
        }

        private static Table BuildSharedPoolTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Shared pools")
                .Border(TableBorder.Rounded)
                .AddColumn("Pool")
                .AddColumn("Group")
                .AddColumn("Cloud PCs in pool")
                .AddColumn("User")
                .AddColumn("UPN");

            var policies = item.FlexPolicies.Where(policy => IsSharedFlexPolicy(policy) && !IsCloudAppsPolicy(policy) && !IsDedicatedFlexPolicy(policy)).ToArray();
            foreach (var policy in policies)
            {
                var cloudPcList = FormatCloudPcList(item.CloudPcs.Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase)));
                foreach (var accessRow in GetPolicyAccessRows(policy, groupMembers))
                {
                    table.AddRow(
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(accessRow.GroupName),
                        Markup.Escape(cloudPcList),
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName));
                }
            }

            if (policies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]No shared Flex pools detected.[/]", "[grey]-[/]", "[grey]-[/]");
            }

            return table;
        }

        private static Table BuildDedicatedMachineTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Dedicated machines")
                .Border(TableBorder.Rounded)
                .AddColumn("User")
                .AddColumn("UPN")
                .AddColumn("Cloud PC")
                .AddColumn("Status")
                .AddColumn("Policy")
                .AddColumn("State");

            var dedicatedPolicies = item.FlexPolicies.Where(IsDedicatedFlexPolicy).ToArray();
            foreach (var policy in dedicatedPolicies)
            {
                var accessRows = GetPolicyAccessRows(policy, groupMembers)
                    .Where(row => row.UserPrincipalName != "-")
                    .OrderBy(row => row.UserName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var cloudPcsForPolicy = item.CloudPcs
                    .Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var accessRow in accessRows)
                {
                    var cloudPc = cloudPcsForPolicy.FirstOrDefault(pc => string.Equals(pc.UserPrincipalName, accessRow.UserPrincipalName, StringComparison.OrdinalIgnoreCase));
                    table.AddRow(
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName),
                        Markup.Escape(cloudPc?.Name ?? "-"),
                        Markup.Escape(cloudPc?.Status ?? "-"),
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(GetDedicatedUserState(cloudPc)));
                }
            }

            if (dedicatedPolicies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No dedicated Flex policy detected.[/]", "[grey]-[/]");
            }

            return table;
        }

        private static string FormatCloudPcList(IEnumerable<CloudPcSummary> cloudPcs)
        {
            var names = cloudPcs.Select(pc => pc.Name).ToArray();
            return names.Length == 0 ? "-" : string.Join(", ", names);
        }

        private static IReadOnlyList<FlexAccessRow> GetPolicyAccessRows(ProvisioningPolicySummary policy, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var rows = new List<FlexAccessRow>();
            for (var index = 0; index < Math.Max(policy.AssignedGroupIds.Count, policy.AssignedGroupNames.Count); index++)
            {
                var groupId = index < policy.AssignedGroupIds.Count ? policy.AssignedGroupIds[index] : null;
                var groupName = index < policy.AssignedGroupNames.Count ? policy.AssignedGroupNames[index] : groupId ?? "-";
                var members = groupId is not null && groupMembers.TryGetValue(groupId, out var resolvedMembers)
                    ? resolvedMembers
                    : [];

                if (members.Count == 0)
                {
                    rows.Add(new FlexAccessRow(groupName, "No direct users found", "-"));
                    continue;
                }

                rows.AddRange(members.Select(member => new FlexAccessRow(groupName, member.Name, member.UserPrincipalName ?? "-")));
            }

            return rows.Count == 0 ? [new FlexAccessRow("-", "No direct users found", "-")] : rows;
        }

        private static string GetFlexCloudPcMode(CloudPcSummary cloudPc)
        {
            if (cloudPc.ProvisioningPolicyName?.Contains("Dedicated", StringComparison.OrdinalIgnoreCase) == true)
            {
                return "Dedicated";
            }

            return "Shared";
        }

        private static string GetDedicatedUserState(CloudPcSummary? cloudPc)
        {
            if (cloudPc is null)
            {
                return "Eligible, no Cloud PC yet";
            }

            var status = cloudPc.Status ?? string.Empty;
            return status.Contains("provisioning", StringComparison.OrdinalIgnoreCase) ||
                status.Contains("pending", StringComparison.OrdinalIgnoreCase)
                    ? "Provisioning"
                    : "Has Cloud PC";
        }

        private static bool IsCloudAppsPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Cloud-Apps", StringComparison.OrdinalIgnoreCase) ||
                policy.DisplayName.Contains("Cloud Apps", StringComparison.OrdinalIgnoreCase) ||
                policy.DisplayName.Contains("CloudApps", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSharedFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Shared", StringComparison.OrdinalIgnoreCase) ||
                (policy.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true && !IsDedicatedFlexPolicy(policy));
        }

        private static bool IsDedicatedFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Dedicated", StringComparison.OrdinalIgnoreCase);
        }

    private async Task<IReadOnlyList<ProvisioningPolicySummary>> LoadProvisioningPoliciesAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading provisioning policies...", async _ => await _session.Graph.GetProvisioningPoliciesAsync());
    }

    private void RenderProvisioningPolicyBrowser(
        IReadOnlyList<ProvisioningPolicySummary> allPolicies,
        IReadOnlyList<ProvisioningPolicySummary> visiblePolicies,
        int selectedIndex,
        string filter,
        ProvisioningPolicySortMode sortMode)
    {
        RenderBreadcrumb("Provisioning", "Policies");
        AnsiConsole.Write(CreateProvisioningPolicySummaryPanel(allPolicies, visiblePolicies, filter));
        AnsiConsole.Write(CreateProvisioningPolicyTable(visiblePolicies, selectedIndex));
        AnsiConsole.MarkupLine($"[grey]Sort: {FormatProvisioningPolicySortMode(sortMode)} | Up/Down move | Enter actions | / filter | C clear | S sort | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    private static Panel CreateProvisioningPolicySummaryPanel(IReadOnlyList<ProvisioningPolicySummary> allPolicies, IReadOnlyList<ProvisioningPolicySummary> visiblePolicies, string filter)
    {
        var typeSummary = string.Join("  ", allPolicies
            .GroupBy(policy => policy.ProvisioningType ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));
        var joinSummary = string.Join("  ", allPolicies
            .GroupBy(policy => policy.DomainJoinTypes ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));

        return new Panel(new Rows(
                new Markup($"[white]Total[/] {allPolicies.Count}   [white]Visible[/] {visiblePolicies.Count}   [white]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
                new Markup($"[white]Types[/] {Markup.Escape(typeSummary)}"),
                new Markup($"[white]Join[/] {Markup.Escape(joinSummary)}")))
            .Header("Policies")
            .Border(BoxBorder.Rounded);
    }

    private static Table CreateProvisioningPolicyTable(IReadOnlyList<ProvisioningPolicySummary> visiblePolicies, int selectedIndex)
    {
        var widths = GetProvisioningPolicyWidths();
        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Name")
            .AddColumn("Type")
            .AddColumn("Image")
            .AddColumn("Join")
            .AddColumn("SSO");

        var showGroups = Console.WindowWidth >= 130;
        if (showGroups)
        {
            table.AddColumn("Groups");
        }

        if (visiblePolicies.Count == 0)
        {
            var cells = new List<string> { "-", "[grey]No policies match the current filter.[/]", "-", "-", "-", "-" };
            if (showGroups) { cells.Add("-"); }
            table.AddRow(cells.ToArray());
            return table;
        }

        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 16));
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visiblePolicies.Count - pageSize));
        var visible = visiblePolicies.Skip(start).Take(pageSize).ToArray();
        for (var index = 0; index < visible.Length; index++)
        {
            var absoluteIndex = start + index;
            var policy = visible[index];
            var selected = absoluteIndex == selectedIndex;
            var row = new List<string>
            {
                selected ? "[black on #58a6ff]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(policy.DisplayName, widths.Name))) : Markup.Escape(Fit(policy.DisplayName, widths.Name)),
                selected ? Selected(Markup.Escape(Fit(policy.ProvisioningType ?? "-", widths.Type))) : Markup.Escape(Fit(policy.ProvisioningType ?? "-", widths.Type)),
                selected ? Selected(Markup.Escape(Fit(policy.ImageDisplayName ?? "-", widths.Image))) : Markup.Escape(Fit(policy.ImageDisplayName ?? "-", widths.Image)),
                selected ? Selected(Markup.Escape(Fit(policy.DomainJoinTypes ?? "-", widths.Join))) : Markup.Escape(Fit(policy.DomainJoinTypes ?? "-", widths.Join)),
                selected ? Selected(Markup.Escape(Fit(FormatBool(policy.EnableSingleSignOn), widths.Sso))) : Markup.Escape(Fit(FormatBool(policy.EnableSingleSignOn), widths.Sso))
            };
            if (showGroups)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(string.Join(", ", policy.AssignedGroupNames), widths.Groups))) : Markup.Escape(Fit(string.Join(", ", policy.AssignedGroupNames), widths.Groups)));
            }

            table.AddRow(row.ToArray());
        }

        return table;
    }

    private static (int Name, int Type, int Image, int Join, int Sso, int Groups) GetProvisioningPolicyWidths()
    {
        var available = Math.Max(95, Console.WindowWidth - 4);
        const int type = 12;
        const int join = 16;
        const int sso = 5;
        var showGroups = Console.WindowWidth >= 130;
        var groups = showGroups ? Math.Max(20, (int)(available * 0.20)) : 0;
        var remaining = Math.Max(45, available - type - join - sso - groups - (showGroups ? 6 : 5));
        var name = Math.Clamp((int)(remaining * 0.55), 28, 44);
        var image = Math.Max(22, remaining - name);
        return (name, type, image, join, sso, groups);
    }

    private async Task ShowProvisioningPolicyDetailsAsync(ProvisioningPolicySummary policy)
    {
        var actions = GetProvisioningPolicyActions(policy);
        var selectedActionIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderProvisioningPolicyDetailLayout(policy, actions, selectedActionIndex);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedActionIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedActionIndex = actions.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    var action = actions[selectedActionIndex];
                    if (action == "Back")
                    {
                        return;
                    }
                    await InvokeProvisioningPolicyActionAsync(policy, action);
                    if (action is "Delete")
                    {
                        return;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }

                    if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private static string[] GetProvisioningPolicyActions(ProvisioningPolicySummary policy)
    {
        var actions = new List<string> { "View Cloud PCs", "Export", "Create copy", "Reprovision policy Cloud PCs" };

        if (IsSharedProvisioningPolicy(policy))
        {
            actions.Add("Reprovision (reserve %)");
            actions.Add("Check reprovision status");
        }

        if (IsSharedByEntraGroupPolicy(policy))
        {
            actions.Add("User experience sync");
        }

        if (policy.AssignedGroupIds.Count > 0)
        {
            actions.Add("Manage group members");
        }

        actions.Add("Delete");
        actions.Add("Back");
        return actions.ToArray();
    }

    private static bool IsSharedProvisioningPolicy(ProvisioningPolicySummary policy)
    {
        return policy.ProvisioningType is not null &&
            new[] { "shared", "sharedByUser", "sharedByEntraGroup" }.Any(value =>
                string.Equals(policy.ProvisioningType, value, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// User settings persistence ("user experience sync") is only available for shared-by-Entra-
    /// group policies — not dedicated or shared-by-user — per Microsoft's documentation.
    /// </summary>
    private static bool IsSharedByEntraGroupPolicy(ProvisioningPolicySummary policy)
    {
        return string.Equals(policy.ProvisioningType, "sharedByEntraGroup", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderProvisioningPolicyDetailLayout(ProvisioningPolicySummary policy, IReadOnlyList<string> actions, int selectedActionIndex)
    {
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName);
        var details = new Panel(new Rows(
                new Markup(PropertyInline("Name", policy.DisplayName)),
                new Markup(PropertyInline("Description", policy.Description ?? "-")),
                new Markup(PropertyInline("Type", policy.ProvisioningType ?? "-")),
                new Markup(PropertyInline("Image", policy.ImageDisplayName ?? "-")),
                new Markup(PropertyInline("Image type", policy.ImageType ?? "-")),
                new Markup(PropertyInline("Domain join", policy.DomainJoinTypes ?? "-")),
                new Markup(PropertyInline("Single sign-on", FormatBool(policy.EnableSingleSignOn))),
                new Markup(PropertyInline("Local admin", FormatBool(policy.LocalAdminEnabled))),
                new Markup(PropertyInline("Naming template", policy.CloudPcNamingTemplate ?? "-")),
                new Markup(PropertyInline("Cloud PC group", policy.CloudPcGroupDisplayName ?? "-")),
                new Markup(PropertyInline("Managed by", policy.ManagedBy ?? "-")),
                new Markup(PropertyInline("Grace period hours", policy.GracePeriodInHours?.ToString() ?? "-")),
                new Markup(PropertyBlock("Assigned groups", string.Join(", ", policy.AssignedGroupNames))),
                new Markup(PropertyBlock("Policy ID", policy.Id))))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
        var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
            .Header("Actions")
            .Border(BoxBorder.Rounded);

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, actionPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(actionPanel);
        }

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back | P or Ctrl+K command palette[/]");
    }

    private async Task InvokeProvisioningPolicyActionAsync(ProvisioningPolicySummary policy, string action)
    {
        switch (action)
        {
            case "View Cloud PCs":
                await ShowCloudPcsForProvisioningPolicyAsync(policy);
                break;
            case "Export":
                await ExportProvisioningPolicyAsync(policy);
                break;
            case "Create copy":
                await CreateProvisioningPolicyCopyAsync(policy);
                break;
            case "Reprovision policy Cloud PCs":
                await ReprovisionProvisioningPolicyCloudPcsAsync(policy);
                break;
            case "Reprovision (reserve %)":
                await ApplyProvisioningPolicyReservePercentageAsync(policy);
                break;
            case "Check reprovision status":
                await ShowProvisioningPolicyApplyStatusAsync(policy);
                break;
            case "User experience sync":
                await ShowUserExperienceSyncAsync(policy);
                break;
            case "Manage group members":
                await ShowProvisioningPolicyGroupMembersAsync(policy);
                break;
            case "Delete":
                await DeleteProvisioningPolicyWithGuardAsync(policy);
                break;
        }
    }

    /// <summary>
    /// Windows 365 requires a provisioning policy to have zero group assignments before it can be
    /// deleted — attempting to delete an assigned policy fails with a 400 Bad Request. Detect that
    /// up front and offer to remove the assignments (via the assign action with an empty list)
    /// before retrying the delete, instead of surfacing a confusing raw error.
    /// </summary>
    private async Task DeleteProvisioningPolicyWithGuardAsync(ProvisioningPolicySummary policy)
    {
        if (policy.AssignedGroupNames.Count > 0)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Delete");
            AnsiConsole.MarkupLine("[yellow]This policy still has group assignments.[/]");
            AnsiConsole.MarkupLine("[grey]Windows 365 requires a provisioning policy to have no assignments before it can be deleted.[/]");
            AnsiConsole.MarkupLine($"[grey]Assigned groups:[/] {Markup.Escape(string.Join(", ", policy.AssignedGroupNames))}");
            AnsiConsole.WriteLine();

            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("How would you like to proceed?")
                    .HighlightStyle(SelectionHighlightStyle())
                    .AddChoices("Remove assignments, then delete", "Cancel"));

            if (choice != "Remove assignments, then delete")
            {
                TimedMessage("[yellow]Delete cancelled.[/]");
                return;
            }

            try
            {
                await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Removing assignments...", async _ => await _session.Graph.UnassignProvisioningPolicyAsync(policy.Id));
            }
            catch (Exception ex)
            {
                if (await HandlePermissionErrorAsync(ex, "Remove assignments", policy.DisplayName) ||
                    HandleLockedResourceError(ex, "Remove assignments", policy.DisplayName))
                {
                    return;
                }

                ShowActionResult("Failed", "Remove assignments", policy.DisplayName, "[red]Failed to remove assignments.[/]", ex.Message);
                return;
            }
        }

        await ConfirmAndRunAsync("Delete", policy.DisplayName, async () => await _session.Graph.DeleteProvisioningPolicyAsync(policy.Id), "Policy", policy.DisplayName);
    }

    private async Task ShowCloudPcsForProvisioningPolicyAsync(ProvisioningPolicySummary policy)
    {
        var cloudPcs = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading policy Cloud PCs...", async _ => await _session.Graph.GetCloudPcsByProvisioningPolicyAsync(policy.Id));
        if (cloudPcs.Count == 0)
        {
            TimedMessage("[yellow]No Cloud PCs were returned for this provisioning policy.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = CloudPcSortMode.Name;
        while (true)
        {
            var visibleCloudPcs = SortCloudPcs(FilterCloudPcs(cloudPcs, filter), sortMode);
            if (visibleCloudPcs.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleCloudPcs.Count)
            {
                selectedIndex = visibleCloudPcs.Count - 1;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Cloud PCs");
            AnsiConsole.Write(CreateCloudPcSummaryPanel(cloudPcs, visibleCloudPcs, filter));
            AnsiConsole.Write(CreateCloudPcTable(cloudPcs, visibleCloudPcs, selectedIndex, filter));
            AnsiConsole.MarkupLine($"[grey]Sort: {FormatCloudPcSortMode(sortMode)} | Enter actions | D disk | N snapshots | Z resize | Y sync | / filter | C clear | S sort | R refresh | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleCloudPcs.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visibleCloudPcs.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visibleCloudPcs.Count - 1);
                    break;
                case ConsoleKey.Enter:
                case ConsoleKey.A:
                    if (visibleCloudPcs.Count > 0)
                    {
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex]);
                    }
                    break;
                case ConsoleKey.R:
                    cloudPcs = await _session.Graph.GetCloudPcsByProvisioningPolicyAsync(policy.Id);
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextCloudPcSortMode(sortMode);
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'd' or 'D')
                    {
                        await ShowDiskSpaceAsync(visibleCloudPcs[selectedIndex]);
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'n' or 'N')
                    {
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex], "Snapshots");
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'z' or 'Z')
                    {
                        await ShowResizeAsync(visibleCloudPcs[selectedIndex]);
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'y' or 'Y')
                    {
                        await InvokeCloudPcActionAsync(visibleCloudPcs[selectedIndex], "Sync");
                    }
                    break;
            }
        }
    }

    private async Task ExportProvisioningPolicyAsync(ProvisioningPolicySummary policy)
    {
        var defaultPath = Path.Combine(Environment.CurrentDirectory, $"{SanitizeFileName(policy.DisplayName)}.json");
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Export");
        var path = AnsiConsole.Prompt(
            new TextPrompt<string>($"Export path [[{Markup.Escape(defaultPath)}]]:")
                .AllowEmpty());
        if (string.IsNullOrWhiteSpace(path))
        {
            path = defaultPath;
        }

        var exportJson = _session.Graph.ExportProvisioningPolicyJson(policy);
        await File.WriteAllTextAsync(path, exportJson);
        ShowActionResult("Exported", "Export", path, "[green]Exported.[/]");
    }

    private async Task CreateProvisioningPolicyCopyAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Create copy");
        var displayName = AnsiConsole.Ask<string>("New policy display name:");
        if (string.IsNullOrWhiteSpace(displayName))
        {
            TimedMessage("[yellow]Create copy cancelled. Display name is required.[/]");
            return;
        }

        var assign = AskYesNo("Recreate assignment targets on the new policy?");
        await ConfirmAndRunAsync(
            "Create copy",
            $"{policy.DisplayName} to {displayName}",
            async () => await _session.Graph.CreateProvisioningPolicyCopyAsync(policy, displayName, assign),
            "Policy",
            policy.DisplayName);
    }

    private async Task CreateProvisioningPolicyWizardAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Create policy");
        AnsiConsole.MarkupLine("[grey]Create a new Windows 365 provisioning policy.[/]");
        AnsiConsole.WriteLine();

        var displayName = AnsiConsole.Ask<string>("Policy display name:");
        if (string.IsNullOrWhiteSpace(displayName))
        {
            TimedMessage("[yellow]Create policy cancelled. Display name is required.[/]");
            return;
        }

        var description = AnsiConsole.Prompt(new TextPrompt<string>("Description [[optional]]:").AllowEmpty());

        var provisioningTypeChoice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Provisioning type")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Dedicated", "Shared by user", "Shared by Entra group", "Back"));
        if (provisioningTypeChoice == "Back")
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        var provisioningType = provisioningTypeChoice switch
        {
            "Dedicated" => "dedicated",
            "Shared by user" => "sharedByUser",
            "Shared by Entra group" => "sharedByEntraGroup",
            _ => "dedicated"
        };

        IReadOnlyList<GraphTableRow> images;
        try
        {
            images = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading gallery images...", async _ => await _session.Graph.GetGalleryImageRowsAsync());
        }
        catch (Exception ex)
        {
            if (!await HandlePermissionErrorAsync(ex, "Create policy", displayName) &&
                !HandleLockedResourceError(ex, "Create policy", displayName))
            {
                ShowActionResult("Failed", "Create policy", displayName, "[red]Failed to load gallery images.[/]", ex.Message);
            }
            return;
        }

        if (images.Count == 0)
        {
            TimedMessage("[yellow]No gallery images are available to select.[/]");
            return;
        }

        // notSupported gallery images can't be used to provision new Cloud PCs and Graph rejects
        // the create with a generic 400 if you pick one — filter them out so only images that
        // can actually be selected are shown.
        var selectableImages = images
            .Where(image => !string.Equals(GetOptionalField(image, "status"), "notSupported", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (selectableImages.Length == 0)
        {
            selectableImages = images.ToArray();
        }

        var imageHeader = Row("Image", 50, "Status", 14, "OS version", 16);
        var selectedImage = SelectFromTable(
            "Select gallery image",
            imageHeader,
            selectableImages,
            image => Row(image.Title, 50, GetField(image, "status"), 14, GetField(image, "osVersionNumber"), 16));
        if (selectedImage is null)
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        var imageId = GetOptionalField(selectedImage, "id", "Id", "ID");
        if (string.IsNullOrWhiteSpace(imageId))
        {
            TimedMessage("[red]Selected image is missing an ID; cannot continue.[/]");
            return;
        }

        var namingTemplate = AnsiConsole.Prompt(
            new TextPrompt<string>("Cloud PC naming template:")
                .DefaultValue("CPC-%USERNAME:5%-%RAND:5%"));

        IReadOnlyList<GraphTableRow> regions;
        try
        {
            regions = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading supported regions...", async _ => await _session.Graph.GetSupportedRegionRowsAsync());
        }
        catch (Exception ex)
        {
            if (!await HandlePermissionErrorAsync(ex, "Create policy", displayName) &&
                !HandleLockedResourceError(ex, "Create policy", displayName))
            {
                ShowActionResult("Failed", "Create policy", displayName, "[red]Failed to load supported regions.[/]", ex.Message);
            }
            return;
        }

        var availableRegions = regions
            .Where(region => string.IsNullOrWhiteSpace(GetOptionalField(region, "supportedSolution")) ||
                string.Equals(GetOptionalField(region, "supportedSolution"), "windows365", StringComparison.OrdinalIgnoreCase))
            .Where(region => string.Equals(GetOptionalField(region, "regionStatus"), "available", StringComparison.OrdinalIgnoreCase))
            .GroupBy(region => region.Title, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        if (availableRegions.Length == 0)
        {
            availableRegions = regions
                .GroupBy(region => region.Title, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
        }

        if (availableRegions.Length == 0)
        {
            TimedMessage("[yellow]No supported regions were returned; a region is required for Microsoft Entra join.[/]");
            return;
        }

        // Microsoft Entra join requires either a region or an on-premises network connection —
        // leaving both empty causes Graph to reject the create with a generic 400, so a region
        // selection here is mandatory rather than optional.
        var regionOptions = availableRegions.Select(region => region.Title).Append("Back").ToArray();
        var regionChoice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Region for Microsoft Entra joined Cloud PCs")
                .HighlightStyle(SelectionHighlightStyle())
                .PageSize(18)
                .AddChoices(regionOptions));

        if (regionChoice == "Back")
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        var regionRow = availableRegions.First(region => region.Title == regionChoice);
        var regionName = regionRow.Title;

        var enableSso = AskYesNo("Enable single sign-on?", defaultToYes: false);
        var localAdmin = AskYesNo("Enable local admin?", defaultToYes: false);

        var assignGroupId = AnsiConsole.Prompt(
            new TextPrompt<string>("Assign to Entra group ID [[optional — paste the group's object ID]]:")
                .AllowEmpty());

        await ConfirmAndRunAsync(
            "Create policy",
            displayName,
            async () => await _session.Graph.CreateProvisioningPolicyAsync(
                displayName,
                description,
                provisioningType,
                imageId,
                selectedImage.Title,
                "gallery",
                "azureADJoin",
                regionName,
                namingTemplate,
                enableSso,
                localAdmin,
                string.IsNullOrWhiteSpace(assignGroupId) ? null : assignGroupId.Trim()),
            "Policy",
            displayName);
    }

    private async Task ReprovisionProvisioningPolicyCloudPcsAsync(ProvisioningPolicySummary policy)
    {
        var osChoice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Policy reprovision OS version")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Keep policy/default", "Windows 11", "Windows 10", "Back"));
        if (osChoice == "Back")
        {
            return;
        }

        var accountChoice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Policy reprovision user account type")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Keep policy/default", "Standard user", "Administrator", "Back"));
        if (accountChoice == "Back")
        {
            return;
        }

        AnsiConsole.WriteLine();
        var excludeInput = AnsiConsole.Prompt(
            new TextPrompt<string>("Exclude Cloud PCs by name, ID, or UPN, comma-separated [[optional]]:")
                .AllowEmpty());
        var exclusions = string.IsNullOrWhiteSpace(excludeInput)
            ? []
            : excludeInput.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var osVersion = osChoice switch
        {
            "Windows 11" => "windows11",
            "Windows 10" => "windows10",
            _ => null
        };
        var accountType = accountChoice switch
        {
            "Standard user" => "standardUser",
            "Administrator" => "administrator",
            _ => null
        };

        await ConfirmAndRunAsync(
            "Reprovision",
            $"{policy.DisplayName} policy Cloud PCs",
            async () => await _session.Graph.ReprovisionCloudPcsByPolicyAsync(policy.Id, osVersion, accountType, exclusions),
            "Policy",
            policy.DisplayName);
    }

    private async Task ApplyProvisioningPolicyReservePercentageAsync(ProvisioningPolicySummary policy)
    {
        var reservePercentage = PromptForInteger(
            $"Reprovision (reserve %) — {policy.DisplayName}",
            "Reprovisions Frontline shared Cloud PCs under this policy while keeping the given percentage available. Cloud PCs are only reprovisioned once they have no active connected user. Enter the percentage to keep available (0-99), or Esc/B/Q to cancel.",
            0,
            99);
        if (reservePercentage is null)
        {
            TimedMessage("[yellow]Reprovision cancelled.[/]");
            return;
        }

        var forceLogoff = reservePercentage == 0 &&
            AskYesNo("Forcibly sign out connected users and reprovision immediately?", defaultToYes: false);

        await ConfirmAndRunAsync(
            "Reprovision (reserve %)",
            $"{policy.DisplayName} — reserve {reservePercentage}%{(forceLogoff ? ", force logoff" : string.Empty)}",
            async () => await _session.Graph.ApplyProvisioningPolicyAsync(policy.Id, reservePercentage.Value, forceLogoff),
            "Policy",
            policy.DisplayName);
    }

    private async Task ShowProvisioningPolicyApplyStatusAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Reprovision status");

        CloudPcPolicyApplyActionResult? result;
        try
        {
            result = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Checking reprovision status...", async _ => await _session.Graph.GetProvisioningPolicyApplyActionResultAsync(policy.Id));
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Check reprovision status", policy.DisplayName))
            {
                return;
            }

            if (HandleLockedResourceError(ex, "Check reprovision status", policy.DisplayName))
            {
                return;
            }

            TimedMessage($"[red]Failed to retrieve reprovision status: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (result is null)
        {
            TimedMessage("[yellow]No reprovision status is available yet.[/]");
            return;
        }

        var statusMarkup = result.Status?.ToLowerInvariant() switch
        {
            "succeeded" => "[green]Succeeded[/]",
            "pending" => "[yellow]Pending[/]",
            "failed" => "[red]Failed[/]",
            _ => Markup.Escape(result.Status ?? "-")
        };

        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Reprovision status");
        AnsiConsole.MarkupLine($"[bold]Status[/] {statusMarkup}");
        AnsiConsole.MarkupLine($"[bold]Started[/] {Markup.Escape(result.StartDateTime?.ToLocalTime().ToString("g") ?? "-")}");
        AnsiConsole.MarkupLine($"[bold]Finished[/] {Markup.Escape(result.FinishDateTime?.ToLocalTime().ToString("g") ?? "-")}");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
    }

    /// <summary>
    /// Admin-level rollup of user experience sync (UES) storage across every shared-by-Entra-group
    /// provisioning policy — total/used/remaining GB and profile count per policy, plus an org-wide
    /// aggregate bar. Enter on a row drills into that policy's per-user profile list.
    /// </summary>
    private async Task ShowUserExperienceSyncOverviewAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var policies = await LoadProvisioningPoliciesAsync();
        var uesPolicies = policies.Where(IsSharedByEntraGroupPolicy).ToArray();

        if (uesPolicies.Length == 0)
        {
            TimedMessage("[yellow]No shared-by-Entra-group provisioning policies were found.[/]");
            return;
        }

        var policiesById = uesPolicies.ToDictionary(policy => policy.Id, StringComparer.OrdinalIgnoreCase);

        async Task<IReadOnlyList<GraphTableRow>> LoadOverviewRowsAsync()
        {
            var rows = new List<GraphTableRow>();

            foreach (var policy in uesPolicies)
            {
                string enabledText;
                string totalText = "-";
                string usedText = "-";
                string remainingText = "-";
                string profileCountText = "-";

                try
                {
                    var context = await _session.Graph.GetUserSettingsPersistenceContextAsync(policy.Id);
                    if (context is null)
                    {
                        enabledText = "Not configured";
                    }
                    else if (!context.Enabled)
                    {
                        enabledText = "Disabled";
                    }
                    else
                    {
                        enabledText = "Enabled";
                        var usage = await _session.Graph.GetUserSettingsPersistenceUsageAsync(context);
                        if (usage is not null)
                        {
                            totalText = (usage.TotalAllocatedStorageInGB ?? 0).ToString("0.#");
                            usedText = (usage.UsedStorageInGB ?? 0).ToString("0.#");
                            remainingText = (usage.RemainingAvailableStorageInGB ?? Math.Max(0, (usage.TotalAllocatedStorageInGB ?? 0) - (usage.UsedStorageInGB ?? 0))).ToString("0.#");
                        }

                        var profiles = await _session.Graph.GetUserSettingsPersistenceProfilesAsync(context);
                        profileCountText = profiles.Count.ToString();
                    }
                }
                catch (Exception ex)
                {
                    enabledText = $"Error: {ex.Message}";
                }

                var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["policyId"] = policy.Id,
                    ["policyName"] = policy.DisplayName,
                    ["status"] = enabledText,
                    ["totalAllocatedStorageInGB"] = totalText,
                    ["usedStorageInGB"] = usedText,
                    ["remainingAvailableStorageInGB"] = remainingText,
                    ["profileCount"] = profileCountText
                };

                rows.Add(new GraphTableRow(policy.DisplayName, enabledText, fields));
            }

            return rows;
        }

        await ShowGraphRowsAsync(
            "User experience sync overview",
            LoadOverviewRowsAsync,
            GetUserSettingsPersistenceOverviewHeader,
            FormatUserSettingsPersistenceOverviewRow,
            enterAction: async row =>
            {
                var policyId = GetOptionalField(row, "policyId");
                if (policyId is not null && policiesById.TryGetValue(policyId, out var policy))
                {
                    await ShowUserExperienceSyncAsync(policy);
                }
            },
            summaryRenderer: RenderUserSettingsPersistenceOverviewSummary);
    }

    private static void RenderUserSettingsPersistenceOverviewSummary(IReadOnlyList<GraphTableRow> rows)
    {
        double total = 0;
        double used = 0;
        var enabledCount = 0;

        foreach (var row in rows)
        {
            if (!string.Equals(GetField(row, "status"), "Enabled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            enabledCount++;
            if (double.TryParse(GetField(row, "totalAllocatedStorageInGB"), out var rowTotal))
            {
                total += rowTotal;
            }

            if (double.TryParse(GetField(row, "usedStorageInGB"), out var rowUsed))
            {
                used += rowUsed;
            }
        }

        AnsiConsole.MarkupLine($"[bold]Policies with UES enabled[/] {enabledCount} of {rows.Count}");

        if (enabledCount == 0)
        {
            return;
        }

        var remaining = Math.Max(0, total - used);
        AnsiConsole.MarkupLine($"[bold]Total allocated[/] {total:0.#} GB    [bold]Used[/] {used:0.#} GB    [bold]Remaining[/] {remaining:0.#} GB");

        const int barWidth = 40;
        var usedSegments = total > 0 ? (int)Math.Round(barWidth * Math.Clamp(used / total, 0, 1)) : 0;
        var freeSegments = Math.Max(0, barWidth - usedSegments);
        var barColor = total > 0 && used / total >= 0.9 ? "red" : total > 0 && used / total >= 0.75 ? "yellow" : "green";
        AnsiConsole.MarkupLine($"[{barColor}]{new string('█', usedSegments)}[/][grey]{new string('░', freeSegments)}[/]");
    }

    private static string GetUserSettingsPersistenceOverviewHeader()
    {
        var widths = GetUserSettingsPersistenceOverviewWidths();
        return Row("Policy", widths.Policy, "Status", widths.Status, "Total (GB)", widths.Total, "Used (GB)", widths.Used, "Remaining (GB)", widths.Remaining, "Profiles", widths.Profiles);
    }

    private static string FormatUserSettingsPersistenceOverviewRow(GraphTableRow row)
    {
        var widths = GetUserSettingsPersistenceOverviewWidths();
        return Row(
            GetField(row, "policyName"), widths.Policy,
            GetField(row, "status"), widths.Status,
            GetField(row, "totalAllocatedStorageInGB"), widths.Total,
            GetField(row, "usedStorageInGB"), widths.Used,
            GetField(row, "remainingAvailableStorageInGB"), widths.Remaining,
            GetField(row, "profileCount"), widths.Profiles);
    }

    private static (int Policy, int Status, int Total, int Used, int Remaining, int Profiles) GetUserSettingsPersistenceOverviewWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 14;
        const int total = 11;
        const int used = 10;
        const int remaining = 15;
        const int profiles = 9;
        var gaps = 5;
        var policy = Math.Max(20, available - status - total - used - remaining - profiles - gaps);
        return (policy, status, total, used, remaining, profiles);
    }

    /// <summary>
    /// "Manage group members" — lists members of the Entra group(s) assigned to this provisioning
    /// policy, with actions to add or remove members directly from the CLI (mirrors what an admin
    /// would otherwise do in the Entra/Azure AD portal's group membership blade).
    /// </summary>
    private async Task ShowProvisioningPolicyGroupMembersAsync(ProvisioningPolicySummary policy)
    {
        if (policy.AssignedGroupIds.Count == 0)
        {
            TimedMessage("[yellow]This policy has no assigned group.[/]");
            return;
        }

        string groupId;
        string groupName;

        if (policy.AssignedGroupIds.Count == 1)
        {
            groupId = policy.AssignedGroupIds[0];
            groupName = policy.AssignedGroupNames.Count > 0 ? policy.AssignedGroupNames[0] : groupId;
        }
        else
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Manage group members");
            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select a group to manage")
                    .HighlightStyle(SelectionHighlightStyle())
                    .AddChoices(policy.AssignedGroupNames));
            var index = policy.AssignedGroupNames.ToList().IndexOf(choice);
            groupId = policy.AssignedGroupIds[Math.Max(0, index)];
            groupName = choice;
        }

        var selectedIndex = 0;

        while (true)
        {
            List<GroupMemberSummary> members;
            try
            {
                members = (await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync($"Loading members of {groupName}...", async _ => await _session.Graph.GetGroupMembersAsync(groupId)))
                    .OrderBy(member => member.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                if (await HandlePermissionErrorAsync(ex, "Manage group members", groupName, "GroupMember.Read.All or Group.ReadWrite.All") ||
                    HandleLockedResourceError(ex, "Manage group members", groupName))
                {
                    return;
                }

                TimedMessage($"[red]Failed to load group members: {Markup.Escape(ex.Message)}[/]");
                return;
            }

            if (selectedIndex >= members.Count)
            {
                selectedIndex = Math.Max(0, members.Count - 1);
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Manage group members");
            AnsiConsole.MarkupLine($"[#58a6ff]Group members[/] [grey]{Markup.Escape(groupName)}[/]  [grey]({members.Count})[/]");
            AnsiConsole.WriteLine();

            var table = new Table().Border(TableBorder.Rounded).AddColumn(" ").AddColumn("Name").AddColumn("UPN");
            if (members.Count == 0)
            {
                table.AddRow(" ", "[grey]No members found.[/]", "-");
            }
            else
            {
                for (var index = 0; index < members.Count; index++)
                {
                    var member = members[index];
                    var selected = index == selectedIndex;
                    table.AddRow(
                        selected ? "[black on #58a6ff]>[/]" : " ",
                        selected ? Selected(Markup.Escape(member.DisplayName ?? "-")) : Markup.Escape(member.DisplayName ?? "-"),
                        selected ? Selected(Markup.Escape(member.UserPrincipalName ?? "-")) : Markup.Escape(member.UserPrincipalName ?? "-"));
                }
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine("[grey]Up/Down select | Enter view/remove | A add member | R refresh | Esc/B/Q back[/]");

            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, members.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, members.Count - 1);
                    break;
                case ConsoleKey.Enter:
                    if (members.Count > 0)
                    {
                        await ShowGroupMemberDetailAsync(groupId, groupName, members[selectedIndex]);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'a' or 'A')
                    {
                        await AddGroupMemberWizardAsync(groupId, groupName);
                    }
                    else if (key.KeyChar is 'r' or 'R')
                    {
                        // Loop reloads members at the top — nothing else to do here.
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// Detail/actions screen for a single group member — Enter never removes directly; the user
    /// must arrow down to "Remove" and press Enter again, avoiding an accidental destructive action.
    /// </summary>
    private async Task ShowGroupMemberDetailAsync(string groupId, string groupName, GroupMemberSummary member)
    {
        var actions = new[] { "Remove from group", "Back" };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", groupName, "Member");
            var details = new Panel(new Rows(
                    new Markup(PropertyInline("Name", member.DisplayName ?? "-")),
                    new Markup(PropertyInline("UPN", member.UserPrincipalName ?? "-")),
                    new Markup(PropertyBlock("User ID", member.Id))))
                .Header("Member details")
                .Border(BoxBorder.Rounded);

            var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
            var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
                .Header("Actions")
                .Border(BoxBorder.Rounded);

            if (Console.WindowWidth >= 120)
            {
                var grid = new Grid();
                grid.AddColumn();
                grid.AddColumn();
                grid.AddRow(details, actionPanel);
                AnsiConsole.Write(grid);
            }
            else
            {
                AnsiConsole.Write(details);
                AnsiConsole.Write(actionPanel);
            }

            AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Enter:
                    if (actions[selectedActionIndex] == "Back")
                    {
                        return;
                    }

                    await ConfirmAndRunAsync(
                        "Remove from group",
                        $"{member.Name} from {groupName}",
                        async () => await _session.Graph.RemoveGroupMemberAsync(groupId, member.Id),
                        resourceType: "Group member",
                        resourceName: member.Name,
                        requiredPermission: "GroupMember.ReadWrite.All or Group.ReadWrite.All");
                    return;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// Search-then-pick wizard for adding a user to the policy's assigned group. Prompts for a
    /// search term (name/UPN/email prefix), shows matches, then adds the selected user via Graph.
    /// </summary>
    private async Task AddGroupMemberWizardAsync(string groupId, string groupName)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", groupName, "Add member");
        AnsiConsole.MarkupLine($"[#58a6ff]Add member[/] [grey]{Markup.Escape(groupName)}[/]");
        AnsiConsole.WriteLine();

        var query = AnsiConsole.Ask<string>("Search by name, UPN, or email:");
        if (string.IsNullOrWhiteSpace(query))
        {
            return;
        }

        List<GroupMemberSummary> matches;
        try
        {
            matches = (await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Searching directory...", async _ => await _session.Graph.SearchUsersAsync(query)))
                .ToList();
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Add member", groupName, "User.Read.All or Directory.Read.All") ||
                HandleLockedResourceError(ex, "Add member", groupName))
            {
                return;
            }

            TimedMessage($"[red]Failed to search directory: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (matches.Count == 0)
        {
            TimedMessage("[yellow]No matching users were found.[/]");
            return;
        }

        var choiceLabels = matches
            .Select(match => $"{match.DisplayName ?? "-"}  <{match.UserPrincipalName ?? "-"}>")
            .Append("Cancel")
            .ToArray();

        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", groupName, "Add member");
        var selected = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("Select a user to add")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices(choiceLabels));

        if (selected == "Cancel")
        {
            return;
        }

        var selectedIndex = Array.IndexOf(choiceLabels, selected);
        var user = matches[selectedIndex];

        await ConfirmAndRunAsync(
            "Add group member",
            $"{user.Name} to {groupName}",
            async () => await _session.Graph.AddGroupMemberAsync(groupId, user.Id),
            resourceType: "Group member",
            resourceName: user.Name,
            requiredPermission: "GroupMember.ReadWrite.All or Group.ReadWrite.All");
    }

    /// <summary>
    /// "User experience sync" — surfaces the user settings persistence storage usage bar and
    /// per-user profile list for a shared-by-Entra-group provisioning policy, matching the view
    /// the Intune portal shows under a shared policy's assignment details.
    /// </summary>
    private async Task ShowUserExperienceSyncAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "User experience sync");

        ProvisioningPolicyUserSettingsPersistenceContext? context;
        try
        {
            context = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Resolving user experience sync configuration...", async _ =>
                    await _session.Graph.GetUserSettingsPersistenceContextAsync(policy.Id));
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "User experience sync", policy.DisplayName))
            {
                return;
            }

            if (HandleLockedResourceError(ex, "User experience sync", policy.DisplayName))
            {
                return;
            }

            TimedMessage($"[red]Failed to resolve user experience sync configuration: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (context is null)
        {
            TimedMessage("[yellow]No assignment with user experience sync (user settings persistence) was found for this policy.[/]");
            return;
        }

        if (!context.Enabled)
        {
            TimedMessage("[yellow]User experience sync is not enabled on this policy's assignment.[/]");
            return;
        }

        CloudPcUserSettingsPersistenceUsageResult? usage = null;

        async Task<IReadOnlyList<GraphTableRow>> LoadProfilesAsync()
        {
            // Refresh usage alongside the profile list (both on initial load and on manual "R"
            // refresh) so the usage bar reflects reality after a profile is deleted.
            try
            {
                usage = await _session.Graph.GetUserSettingsPersistenceUsageAsync(context);
            }
            catch
            {
                usage = null;
            }

            return await _session.Graph.GetUserSettingsPersistenceProfilesAsync(context);
        }

        await ShowGraphRowsAsync(
            "User profiles",
            LoadProfilesAsync,
            GetUserSettingsPersistenceProfilesHeader,
            FormatUserSettingsPersistenceProfileRow,
            enterAction: row => ShowUserSettingsPersistenceProfileDetailAsync(context, row),
            summaryRenderer: _ =>
            {
                if (usage is not null)
                {
                    RenderUserSettingsPersistenceUsageBar(usage);
                }
                else
                {
                    AnsiConsole.MarkupLine("[yellow]Storage usage is not available.[/]");
                }
            });
    }

    /// <summary>
    /// Detail/actions screen for a single UES profile row (opened via Enter on the profile list).
    /// Shows the profile's fields and offers "Delete" as an explicit, arrow-key-selected action —
    /// Enter on the row itself never deletes directly, avoiding an accidental destructive action.
    /// </summary>
    private async Task ShowUserSettingsPersistenceProfileDetailAsync(ProvisioningPolicyUserSettingsPersistenceContext context, GraphTableRow row)
    {
        var actions = new[] { "Delete", "Back" };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderUserSettingsPersistenceProfileDetailLayout(row, actions, selectedActionIndex);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedActionIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedActionIndex = actions.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    var action = actions[selectedActionIndex];
                    if (action == "Back")
                    {
                        return;
                    }

                    await DeleteUserSettingsPersistenceProfileAsync(context, row);
                    return;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private static void RenderUserSettingsPersistenceProfileDetailLayout(GraphTableRow row, IReadOnlyList<string> actions, int selectedActionIndex)
    {
        RenderBreadcrumb("Provisioning", "Policies", "User experience sync", "Profile");
        var details = new Panel(new Rows(
                new Markup(PropertyInline("User", GetField(row, "userPrincipalName"))),
                new Markup(PropertyInline("Size (GB)", GetField(row, "profileSizeInGB"))),
                new Markup(PropertyInline("Status", GetField(row, "status"))),
                new Markup(PropertyInline("Last attached", GetField(row, "lastProfileAttachedDateTime"))),
                new Markup(PropertyBlock("Profile ID", GetField(row, "profileId")))))
            .Header("Profile details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
        var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
            .Header("Actions")
            .Border(BoxBorder.Rounded);

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, actionPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(actionPanel);
        }

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back[/]");
    }

    /// <summary>
    /// Deletes ("cleans up") a UES disk/profile for one user, after the "Delete" action is chosen
    /// on the profile detail screen. Mirrors the Intune portal's per-profile delete, which puts the
    /// profile into a "deleting" status the next time the list is refreshed (press R).
    /// </summary>
    private async Task DeleteUserSettingsPersistenceProfileAsync(ProvisioningPolicyUserSettingsPersistenceContext context, GraphTableRow row)
    {
        var profileId = GetOptionalField(row, "profileId", "ProfileId");
        var upn = GetOptionalField(row, "userPrincipalName", "UserPrincipalName") ?? "this user";

        if (string.IsNullOrWhiteSpace(profileId))
        {
            TimedMessage("[yellow]This profile has no profileId — it cannot be deleted.[/]");
            return;
        }

        var deleted = false;
        await ConfirmAndRunAsync(
            "Delete user experience sync profile",
            upn,
            async () =>
            {
                await _session.Graph.BatchCleanupUserSettingsPersistenceProfileAsync(context, profileId);
                deleted = true;
            },
            resourceType: "UES profile",
            resourceName: upn);

        if (deleted)
        {
            TimedMessage("[grey]Deletion queued. Press R on the profile list to refresh status.[/]");
        }
    }

    private static void RenderUserSettingsPersistenceUsageBar(CloudPcUserSettingsPersistenceUsageResult usage)
    {
        var total = usage.TotalAllocatedStorageInGB ?? 0;
        var used = usage.UsedStorageInGB ?? 0;
        var remaining = usage.RemainingAvailableStorageInGB ?? Math.Max(0, total - used);

        AnsiConsole.MarkupLine($"[bold]Total allocated[/] {total:0.#} GB    [bold]Used[/] {used:0.#} GB    [bold]Remaining[/] {remaining:0.#} GB");

        const int barWidth = 40;
        var usedSegments = total > 0 ? (int)Math.Round(barWidth * Math.Clamp(used / total, 0, 1)) : 0;
        var freeSegments = Math.Max(0, barWidth - usedSegments);
        var barColor = total > 0 && used / total >= 0.9 ? "red" : total > 0 && used / total >= 0.75 ? "yellow" : "green";
        AnsiConsole.MarkupLine($"[{barColor}]{new string('█', usedSegments)}[/][grey]{new string('░', freeSegments)}[/]");
    }

    private static string GetUserSettingsPersistenceProfilesHeader()
    {
        var widths = GetUserSettingsPersistenceProfileWidths();
        return Row("UPN", widths.Upn, "Size (GB)", widths.Size, "Status", widths.Status, "Last attached", widths.LastAttached);
    }

    private static string FormatUserSettingsPersistenceProfileRow(GraphTableRow row)
    {
        var widths = GetUserSettingsPersistenceProfileWidths();
        return Row(
            GetField(row, "userPrincipalName"), widths.Upn,
            GetField(row, "profileSizeInGB"), widths.Size,
            GetField(row, "status"), widths.Status,
            GetField(row, "lastProfileAttachedDateTime"), widths.LastAttached);
    }

    private static (int Upn, int Size, int Status, int LastAttached) GetUserSettingsPersistenceProfileWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int size = 10;
        const int status = 14;
        const int lastAttached = 20;
        var gaps = 3;
        var upn = Math.Max(20, available - size - status - lastAttached - gaps);
        return (upn, size, status, lastAttached);
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(character => invalid.Contains(character) ? '-' : character));
    }

    private static string FormatBool(bool? value)
    {
        return value is null ? "-" : value.Value ? "Yes" : "No";
    }

    private async Task ShowActionHistoryAsync()
    {
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderActionHistory(selectedIndex);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, ActionHistory.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, ActionHistory.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.C:
                    ActionHistory.Clear();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Enter:
                    if (ActionHistory.Count > 0)
                    {
                        await OpenRemoteActionsFromHistoryAsync(ActionHistory[selectedIndex]);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private static void RenderActionHistory(int selectedIndex)
    {
        RenderBreadcrumb("Action history");
        var submitted = ActionHistory.Count(item => item.Status.Equals("Submitted", StringComparison.OrdinalIgnoreCase));
        var failed = ActionHistory.Count(item => item.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase));
        AnsiConsole.MarkupLine("[#58a6ff]Action history[/]");
        AnsiConsole.MarkupLine($"[grey]Total: {ActionHistory.Count} | Submitted: {submitted} | Failed: {failed}[/]");
        AnsiConsole.WriteLine();

        if (ActionHistory.Count == 0)
        {
            AnsiConsole.MarkupLine("[grey]No actions have been submitted in this session.[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]H opens this log from other screens. Esc/B/Q back.[/]");
            return;
        }

        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Time")
            .AddColumn("Result")
            .AddColumn("Action")
            .AddColumn("Type")
            .AddColumn("Resource")
            .AddColumn("Target");
        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 18));
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, ActionHistory.Count - pageSize));
        var visible = ActionHistory.Skip(start).Take(pageSize).ToArray();
        foreach (var item in visible.Select((value, index) => new { value, index }))
        {
            var absoluteIndex = start + item.index;
            var selectedMarker = absoluteIndex == selectedIndex ? "[black on #58a6ff]>[/]" : " ";
            table.AddRow(
                selectedMarker,
                Markup.Escape(item.value.RequestedAt.ToLocalTime().ToString("t")),
                ActionStatusCell(item.value.Status),
                Markup.Escape(item.value.Action),
                Markup.Escape(item.value.ResourceType),
                Markup.Escape(item.value.ResourceName ?? "-"),
                Markup.Escape(item.value.Target));
        }

        AnsiConsole.Write(table);

        var selected = ActionHistory[Math.Min(selectedIndex, ActionHistory.Count - 1)];
        if (!string.IsNullOrWhiteSpace(selected.Detail))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine($"[grey]Selected detail: {Markup.Escape(Fit(selected.Detail, Math.Max(40, Console.WindowWidth - 20)))}[/]");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter remote actions for Cloud PC rows | C clear | Esc/B/Q back[/]");
    }

    private static string ActionStatusCell(string status)
    {
        return status.Equals("Failed", StringComparison.OrdinalIgnoreCase)
            ? $"[red]{Markup.Escape(status)}[/]"
            : status.Equals("Submitted", StringComparison.OrdinalIgnoreCase)
                ? $"[green]{Markup.Escape(status)}[/]"
                : Markup.Escape(status);
    }

    private async Task OpenRemoteActionsFromHistoryAsync(ActionHistoryItem item)
    {
        if (!item.ResourceType.Equals("Cloud PC", StringComparison.OrdinalIgnoreCase))
        {
            TimedMessage("[yellow]Remote action history is only available for Cloud PC rows.[/]");
            return;
        }

        var cloudPcs = await LoadCloudPcsAsync();
        var cloudPc = cloudPcs.FirstOrDefault(pc =>
            string.Equals(pc.Name, item.ResourceName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.ManagedDeviceName, item.ResourceName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.DisplayName, item.ResourceName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.Id, item.Target, StringComparison.OrdinalIgnoreCase));

        if (cloudPc is null)
        {
            TimedMessage("[yellow]Could not resolve this action to a Cloud PC.[/]");
            return;
        }

        await ShowCloudPcDetailsAsync(cloudPc, "Remote action history");
    }

    private async Task ShowReportsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        while (true)
        {
            AnsiConsole.Clear();
            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("[#58a6ff]Reports[/]")
                    .HighlightStyle(SelectionHighlightStyle())
                    .AddChoices(
                    "Sign-in status",
                        "Connectivity history",
                        "Launch details",
                        "Cloud PC reports",
                        "Back"));

            switch (choice)
            {
                case "Sign-in status":
                    await ShowGraphRowsAsync(
                        "Windows 365 Cloud PC sign-in status",
                        async () => await _session.Graph.GetSignInStatusRowsAsync(),
                        GetUsageReportHeader,
                        FormatUsageReportRow,
                        OpenCloudPcFromReportRowAsync);
                    break;
                case "Connectivity history":
                    await ShowConnectivityHistoryAsync();
                    break;
                case "Launch details":
                    await ShowGraphRowsAsync(
                        "Windows 365 launch details",
                        async () => await _session.Graph.GetLaunchDetailRowsAsync(),
                        GetLaunchDetailsHeader,
                        FormatLaunchDetailsRow);
                    break;
                case "Cloud PC reports":
                    await ShowCloudPcReportsAsync();
                    break;
                case "Back":
                    return;
            }
        }
    }

    private async Task ShowTenantSettingsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        while (true)
        {
            AnsiConsole.Clear();
            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("[#58a6ff]Tenant settings[/]")
                    .HighlightStyle(SelectionHighlightStyle())
                    .AddChoices(
                        "Organization settings",
                        "Setting profiles",
                        "User settings",
                        "Back"));

            switch (choice)
            {
                case "Organization settings":
                    await ShowGraphRowsAsync(
                        "Windows 365 organization settings",
                        async () => await _session.Graph.GetOrganizationSettingsAsync(),
                        GetOrganizationSettingsHeader,
                        FormatOrganizationSettingRow);
                    break;
                case "Setting profiles":
                    await ShowGraphRowsAsync(
                        "Windows 365 setting profiles",
                        async () => await _session.Graph.GetSettingProfilesAsync(),
                        GetSettingProfilesHeader,
                        FormatSettingProfileRow);
                    break;
                case "User settings":
                    await ShowGraphRowsAsync(
                        "Windows 365 user settings",
                        async () => await _session.Graph.GetUserSettingsAsync(),
                        GetUserSettingsHeader,
                        FormatUserSettingRow);
                    break;
                case "Back":
                    return;
            }
        }
    }

    private async Task ShowCatalogAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var choices = new[] { "Service plans", "Gallery images", "Custom images", "Supported regions", "Back" };
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Catalog");
            AnsiConsole.MarkupLine("[#58a6ff]Catalog[/]");
            AnsiConsole.MarkupLine("[grey]Plans, images, and regions used by Windows 365.[/]");
            AnsiConsole.WriteLine();

            for (var index = 0; index < choices.Length; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on #58a6ff]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter select | Esc/B/Q back[/]");
            RenderStatusBar();
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(choices.Length - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = choices.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    switch (choices[selectedIndex])
                    {
                        case "Service plans":
                            await ShowGraphRowsAsync("Windows 365 service plans", _session.Graph.GetServicePlanRowsAsync, GetServicePlansHeader, FormatServicePlanRow);
                            break;
                        case "Gallery images":
                            await ShowGraphRowsAsync("Windows 365 gallery images", _session.Graph.GetGalleryImageRowsAsync, GetGalleryImagesHeader, FormatGalleryImageRow);
                            break;
                        case "Custom images":
                            await ShowGraphRowsAsync("Windows 365 custom images", _session.Graph.GetCustomImageRowsAsync, GetCustomImagesHeader, FormatCustomImageRow);
                            break;
                        case "Supported regions":
                            await ShowGraphRowsAsync("Windows 365 supported regions", _session.Graph.GetSupportedRegionRowsAsync, GetSupportedRegionsHeader, FormatSupportedRegionRow);
                            break;
                        case "Back":
                            return;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private async Task ShowConnectivityHistoryAsync()
    {
        var cloudPcs = await LoadCloudPcsAsync();
        if (cloudPcs.Count == 0)
        {
            TimedMessage("[yellow]No Cloud PCs returned.[/]");
            return;
        }

        while (true)
        {
            AnsiConsole.Clear();
            var cloudPc = SelectCloudPcForConnectivityHistory(cloudPcs);
            if (cloudPc is null)
            {
                return;
            }

            await ShowGraphRowsAsync(
                $"Connectivity history for {cloudPc.Name}",
                async () => await _session.Graph.GetConnectivityHistoryAsync(cloudPc),
                GetConnectivityHistoryHeader,
                FormatConnectivityHistoryRow,
                async _ => await ShowCloudPcDetailsAsync(cloudPc),
                RenderConnectivityHistorySummary);
        }
    }

    private async Task ShowCloudPcReportsAsync()
    {
        var reportNames = new[]
        {
            "dailyAggregatedRemoteConnectionReports",
            "totalAggregatedRemoteConnectionReports",
            "frontlineLicenseUsageReport",
            "frontlineLicenseUsageRealTimeReport",
            "frontlineLicenseHourlyUsageReport",
            "frontlineRealtimeUserConnectionsReport",
            "inaccessibleCloudPcReports",
            "actionStatusReport",
            "performanceTrendReport",
            "regionalConnectionQualityTrendReport",
            "cloudPcUsageCategoryReport"
        };

        while (true)
        {
            var reportName = SelectCloudPcReportName(reportNames);
            if (reportName is null)
            {
                return;
            }

            var top = PromptTopRows();
            if (top is null)
            {
                continue;
            }

            await ShowGraphRowsAsync(
                $"Report: {reportName}",
                async () => await _session.Graph.GetCloudPcReportRowsAsync(reportName, top.Value),
                enterAction: OpenCloudPcFromReportRowAsync);
        }
    }

    private static string? SelectCloudPcReportName(IReadOnlyList<string> reportNames)
    {
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Reports", "Cloud PC reports");
            AnsiConsole.MarkupLine("[#58a6ff]Cloud PC report[/]");
            AnsiConsole.WriteLine();

            var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
            var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, reportNames.Count - pageSize));
            var visible = reportNames.Skip(start).Take(pageSize).ToArray();
            for (var index = 0; index < visible.Length; index++)
            {
                var absoluteIndex = start + index;
                var escaped = Markup.Escape(visible[index]);
                AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                    ? $"[black on #58a6ff]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter select | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(reportNames.Count - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(reportNames.Count - 1, selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = reportNames.Count - 1;
                    break;
                case ConsoleKey.Enter:
                    return reportNames[selectedIndex];
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return null;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    return null;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return null;
                    }
                    break;
            }
        }
    }

    private static int? PromptTopRows()
    {
        return PromptForInteger(
            "Top rows",
            "Enter a positive number, press Enter for 50, or Esc/B/Q to go back.",
            1,
            int.MaxValue,
            50);
    }

    /// <summary>
    /// Standard bounded-integer input prompt used across the app (row counts, percentages, etc.),
    /// so every "type a number" screen behaves identically: type digits, Backspace to edit, Enter
    /// to accept (falling back to defaultValue when left blank, if one is provided), Esc/B/Q to
    /// cancel back to the caller.
    /// </summary>
    private static int? PromptForInteger(string title, string instructions, int min, int max, int? defaultValue = null)
    {
        var input = string.Empty;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb(title);
            AnsiConsole.MarkupLine($"[{AccentColor}]{Markup.Escape(title)}[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(instructions)}[/]");
            AnsiConsole.WriteLine();
            var prompt = defaultValue is null ? $"{title}: " : $"{title} [{defaultValue}]: ";
            AnsiConsole.Markup($"{Markup.Escape(prompt)}{Markup.Escape(input)}");

            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.Enter:
                    if (string.IsNullOrWhiteSpace(input))
                    {
                        if (defaultValue is not null)
                        {
                            return defaultValue;
                        }

                        TimedMessage($"[yellow]Enter a value between {min} and {max}.[/]");
                        break;
                    }

                    if (int.TryParse(input, out var value) && value >= min && value <= max)
                    {
                        return value;
                    }

                    TimedMessage($"[yellow]Enter a value between {min} and {max}.[/]");
                    break;
                case ConsoleKey.Backspace:
                    if (input.Length > 0)
                    {
                        input = input[..^1];
                    }
                    break;
                case ConsoleKey.Escape:
                    return null;
                default:
                    if (string.IsNullOrWhiteSpace(input) && key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return null;
                    }

                    if (char.IsDigit(key.KeyChar) && input.Length < 10)
                    {
                        input += key.KeyChar;
                    }
                    break;
            }
        }
    }

    private async Task ShowGraphRowsAsync(
        string title,
        Func<Task<IReadOnlyList<GraphTableRow>>> loader,
        Func<string>? headerFactory = null,
        Func<GraphTableRow, string>? rowFactory = null,
        Func<GraphTableRow, Task>? enterAction = null,
        Action<IReadOnlyList<GraphTableRow>>? summaryRenderer = null,
        Func<int, int, Task<IReadOnlyList<GraphTableRow>>>? loadMoreAsync = null,
        int pageBatchSize = 0)
    {
        List<GraphTableRow> rows;
        try
        {
            var initial = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Loading {title}...", async _ => await loader());
            rows = initial.ToList();
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, $"Load {title}", title))
            {
                return;
            }

            if (HandleLockedResourceError(ex, $"Load {title}", title))
            {
                return;
            }

            AnsiConsole.Clear();
            AnsiConsole.MarkupLine($"[red]Failed to load {Markup.Escape(title)}.[/]");
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            TimedMessage("[grey]Returning...[/]");
            return;
        }

        if (rows.Count == 0)
        {
            TimedMessage("[yellow]No rows returned.[/]");
            return;
        }

        headerFactory ??= GetDefaultGraphRowsHeader;
        rowFactory ??= FormatDefaultGraphRow;
        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = GraphRowSortMode.None;
        var hasMore = loadMoreAsync is not null && pageBatchSize > 0;

        async Task<bool> TryLoadMoreAsync()
        {
            if (loadMoreAsync is null || !hasMore)
            {
                return false;
            }

            IReadOnlyList<GraphTableRow> more;
            try
            {
                more = await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Loading more rows...", async _ => await loadMoreAsync(rows.Count, pageBatchSize));
            }
            catch (Exception ex)
            {
                hasMore = false;
                TimedMessage($"[yellow]Failed to load more rows: {Markup.Escape(ex.Message)}[/]");
                return false;
            }

            if (more.Count > 0)
            {
                rows.AddRange(more);
            }

            hasMore = more.Count >= pageBatchSize;
            return more.Count > 0;
        }

        while (true)
        {
            var visibleRows = SortGraphRows(FilterGraphRows(rows, filter), sortMode);
            if (visibleRows.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleRows.Count)
            {
                selectedIndex = visibleRows.Count - 1;
            }

            AnsiConsole.Clear();
            RenderGraphRows(title, rows, visibleRows, selectedIndex, filter, sortMode, headerFactory, rowFactory, summaryRenderer, hasMore);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    if (selectedIndex >= visibleRows.Count - 1 && string.IsNullOrEmpty(filter) && hasMore)
                    {
                        if (await TryLoadMoreAsync())
                        {
                            visibleRows = SortGraphRows(FilterGraphRows(rows, filter), sortMode);
                        }
                    }
                    selectedIndex = Math.Min(Math.Max(0, visibleRows.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    if (selectedIndex + 10 >= visibleRows.Count - 1 && string.IsNullOrEmpty(filter) && hasMore)
                    {
                        if (await TryLoadMoreAsync())
                        {
                            visibleRows = SortGraphRows(FilterGraphRows(rows, filter), sortMode);
                        }
                    }
                    selectedIndex = Math.Min(Math.Max(0, visibleRows.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visibleRows.Count - 1);
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextSortMode(sortMode);
                    selectedIndex = 0;
                    break;
                case ConsoleKey.R:
                    try
                    {
                        var refreshed = await AnsiConsole.Status()
                            .Spinner(Spinner.Known.Dots)
                            .StartAsync($"Refreshing {title}...", async _ => await loader());
                        rows = refreshed.ToList();
                        hasMore = loadMoreAsync is not null && pageBatchSize > 0;
                        selectedIndex = 0;
                    }
                    catch (Exception ex)
                    {
                        if (!await HandlePermissionErrorAsync(ex, $"Refresh {title}", title) &&
                            !HandleLockedResourceError(ex, $"Refresh {title}", title))
                        {
                            TimedMessage($"[red]Failed to refresh: {Markup.Escape(ex.Message)}[/]");
                        }
                    }
                    break;
                case ConsoleKey.Enter:
                case ConsoleKey.RightArrow:
                    if (visibleRows.Count == 0)
                    {
                        break;
                    }

                    if (enterAction is null)
                    {
                        ShowGraphRowDetails(title, visibleRows[selectedIndex]);
                    }
                    else
                    {
                        await enterAction(visibleRows[selectedIndex]);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private async Task OpenCloudPcFromReportRowAsync(GraphTableRow row)
    {
        var cloudPc = await ResolveCloudPcFromReportRowAsync(row);
        if (cloudPc is null)
        {
            ShowGraphRowDetails("Report row details", row);
            return;
        }

        await ShowCloudPcDetailsAsync(cloudPc);
    }

    private async Task<CloudPcSummary?> ResolveCloudPcFromReportRowAsync(GraphTableRow row)
    {
        var cloudPcs = await LoadCloudPcsAsync();
        var cloudPcId = GetOptionalField(row, "Cloud PC ID", "CloudPcId", "cloudPcId", "Cloud PC Id");
        if (!string.IsNullOrWhiteSpace(cloudPcId))
        {
            var idMatch = cloudPcs.FirstOrDefault(pc => string.Equals(pc.Id, cloudPcId, StringComparison.OrdinalIgnoreCase));
            if (idMatch is not null)
            {
                return idMatch;
            }
        }

        var cloudPcName = GetOptionalField(row, "Cloud PC", "CloudPcName", "cloudPcName", "ManagedDeviceName", "managedDeviceName", "DisplayName", "displayName");
        return string.IsNullOrWhiteSpace(cloudPcName)
            ? null
            : cloudPcs.FirstOrDefault(pc =>
                string.Equals(pc.Name, cloudPcName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pc.ManagedDeviceName, cloudPcName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pc.DisplayName, cloudPcName, StringComparison.OrdinalIgnoreCase));
    }

    private static CloudPcSummary? SelectCloudPcForConnectivityHistory(IReadOnlyList<CloudPcSummary> cloudPcs)
    {
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Reports", "Connectivity history");
            AnsiConsole.MarkupLine("[#58a6ff]Select Cloud PC for connectivity history[/]");
            var widths = GetConnectivityCloudPcWidths();
            var header = widths.ServicePlan > 0
                ? Row("Name", widths.Name, "Status", widths.Status, "User", widths.User, "Service plan", widths.ServicePlan)
                : Row("Name", widths.Name, "Status", widths.Status, "User", widths.User);
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

            var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
            var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, cloudPcs.Count - pageSize));
            var visible = cloudPcs.Skip(start).Take(pageSize).ToArray();
            for (var index = 0; index < visible.Length; index++)
            {
                var pc = visible[index];
                var absoluteIndex = start + index;
                var row = widths.ServicePlan > 0
                    ? Row(pc.Name, widths.Name, pc.Status ?? "-", widths.Status, pc.UserPrincipalName ?? "-", widths.User, pc.ServicePlanName ?? "-", widths.ServicePlan)
                    : Row(pc.Name, widths.Name, pc.Status ?? "-", widths.Status, pc.UserPrincipalName ?? "-", widths.User);
                var escaped = Markup.Escape(row);
                AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                    ? $"[black on #58a6ff]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter select | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(cloudPcs.Count - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(cloudPcs.Count - 1, selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = cloudPcs.Count - 1;
                    break;
                case ConsoleKey.Enter:
                    return cloudPcs[selectedIndex];
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return null;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return null;
                    }
                    break;
            }
        }
    }

    private static (int Name, int Status, int User, int ServicePlan) GetConnectivityCloudPcWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 12;
        var showServicePlan = available >= 112;
        var gaps = showServicePlan ? 3 : 2;
        var remaining = Math.Max(42, available - status - gaps);
        var servicePlan = showServicePlan ? Math.Max(24, (int)(remaining * 0.28)) : 0;
        var user = Math.Max(24, (int)((remaining - servicePlan) * 0.48));
        var name = Math.Max(24, remaining - servicePlan - user);
        return (name, status, user, servicePlan);
    }

    private static void RenderGraphRows(
        string title,
        IReadOnlyList<GraphTableRow> allRows,
        IReadOnlyList<GraphTableRow> visibleRows,
        int selectedIndex,
        string filter,
        GraphRowSortMode sortMode,
        Func<string> headerFactory,
        Func<GraphTableRow, string> rowFactory,
        Action<IReadOnlyList<GraphTableRow>>? summaryRenderer,
        bool hasMore = false)
    {
        var header = headerFactory();
        RenderBreadcrumb(title);
        AnsiConsole.MarkupLine($"[#58a6ff]{Markup.Escape(title)}[/]");
        var rowCountText = hasMore ? $"{allRows.Count}+" : allRows.Count.ToString();
        AnsiConsole.MarkupLine($"[grey]Rows: {rowCountText} | Visible: {visibleRows.Count} | Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)} | Sort: {FormatSortMode(sortMode)}[/]");
        AnsiConsole.WriteLine();
        summaryRenderer?.Invoke(allRows);
        if (summaryRenderer is not null)
        {
            AnsiConsole.WriteLine();
        }
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

        var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
        if (visibleRows.Count == 0)
        {
            AnsiConsole.MarkupLine("[grey]No rows match the current filter.[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]/ or F filter | C clear | S sort | R refresh | Esc/B/Q back[/]");
            RenderStatusBar();
            return;
        }

        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visibleRows.Count - pageSize));
        var visible = visibleRows.Skip(start).Take(pageSize).ToArray();

        for (var index = 0; index < visible.Length; index++)
        {
            var absoluteIndex = start + index;
            var escaped = Markup.Escape(rowFactory(visible[index]));
            AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                ? $"[black on #58a6ff]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        var hint = hasMore
            ? "Up/Down move (loads more at end) | PgUp/PgDn page | Enter/-> details | / or F filter | C clear | S sort | R refresh | Esc/B/Q back"
            : "Up/Down move | PgUp/PgDn page | Enter/-> details | / or F filter | C clear | S sort | R refresh | Esc/B/Q back";
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(hint)}[/]");
        RenderStatusBar();
    }

    private static string GetDefaultGraphRowsHeader()
    {
        var widths = GetDefaultGraphRowsWidths();
        return Row("Name", widths.Name, "Summary", widths.Summary);
    }

    private static (int Name, int Summary) GetDefaultGraphRowsWidths()
    {
        var available = Math.Max(70, Console.WindowWidth - 4);
        var name = Math.Max(28, Math.Min(42, available / 3));
        var summary = Math.Max(28, available - name - 1);
        return (name, summary);
    }

    private static string FormatDefaultGraphRow(GraphTableRow row)
    {
        var widths = GetDefaultGraphRowsWidths();
        return Row(row.Title, widths.Name, row.Summary, widths.Summary);
    }

    private static IReadOnlyList<GraphTableRow> FilterGraphRows(IReadOnlyList<GraphTableRow> rows, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return rows;
        }

        return rows
            .Where(row =>
                Contains(row.Title, filter) ||
                Contains(row.Summary, filter) ||
                row.Fields.Any(field => Contains(field.Key, filter) || Contains(field.Value, filter)))
            .ToArray();
    }

    private static IReadOnlyList<GraphTableRow> SortGraphRows(IReadOnlyList<GraphTableRow> rows, GraphRowSortMode sortMode)
    {
        return sortMode switch
        {
            GraphRowSortMode.TitleAscending => rows.OrderBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray(),
            GraphRowSortMode.TitleDescending => rows.OrderByDescending(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray(),
            GraphRowSortMode.SummaryAscending => rows.OrderBy(row => row.Summary, StringComparer.OrdinalIgnoreCase).ThenBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray(),
            GraphRowSortMode.SummaryDescending => rows.OrderByDescending(row => row.Summary, StringComparer.OrdinalIgnoreCase).ThenBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => rows
        };
    }

    private static GraphRowSortMode NextSortMode(GraphRowSortMode sortMode)
    {
        return sortMode switch
        {
            GraphRowSortMode.None => GraphRowSortMode.TitleAscending,
            GraphRowSortMode.TitleAscending => GraphRowSortMode.TitleDescending,
            GraphRowSortMode.TitleDescending => GraphRowSortMode.SummaryAscending,
            GraphRowSortMode.SummaryAscending => GraphRowSortMode.SummaryDescending,
            _ => GraphRowSortMode.None
        };
    }

    private static string FormatSortMode(GraphRowSortMode sortMode)
    {
        return sortMode switch
        {
            GraphRowSortMode.TitleAscending => "title asc",
            GraphRowSortMode.TitleDescending => "title desc",
            GraphRowSortMode.SummaryAscending => "summary asc",
            GraphRowSortMode.SummaryDescending => "summary desc",
            _ => "none"
        };
    }

    private static string GetUsageReportHeader()
    {
        var widths = GetUsageReportWidths();
        return widths.ServicePlan > 0
            ? Row("Cloud PC", widths.CloudPc, "Sign-in", widths.Status, "Days", widths.Power, "Last active", widths.User, "Service plan", widths.ServicePlan)
            : Row("Cloud PC", widths.CloudPc, "Sign-in", widths.Status, "Days", widths.Power, "Last active", widths.User);
    }

    private static string FormatUsageReportRow(GraphTableRow row)
    {
        var widths = GetUsageReportWidths();
        return widths.ServicePlan > 0
            ? Row(
                GetField(row, "Cloud PC"), widths.CloudPc,
                GetField(row, "SignInStatus"), widths.Status,
                GetField(row, "DaysSinceLastSignIn"), widths.Power,
                GetField(row, "LastActiveTime"), widths.User,
                GetField(row, "Service plan"), widths.ServicePlan)
            : Row(
                GetField(row, "Cloud PC"), widths.CloudPc,
                GetField(row, "SignInStatus"), widths.Status,
                GetField(row, "DaysSinceLastSignIn"), widths.Power,
                GetField(row, "LastActiveTime"), widths.User);
    }

    private static (int CloudPc, int Status, int Power, int User, int ServicePlan) GetUsageReportWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 14;
        const int power = 8;
        var showServicePlan = available >= 118;
        var gaps = showServicePlan ? 4 : 3;
        var remaining = Math.Max(42, available - status - power - gaps);
        var servicePlan = showServicePlan ? Math.Max(22, (int)(remaining * 0.28)) : 0;
        var user = Math.Max(22, (int)((remaining - servicePlan) * 0.48));
        var cloudPc = Math.Max(24, remaining - servicePlan - user);
        return (cloudPc, status, power, user, servicePlan);
    }

    private static string GetConnectivityHistoryHeader()
    {
        var widths = GetConnectivityHistoryWidths();
        return widths.Event > 0
            ? Row("Time", widths.Time, "Type", widths.Type, "Event", widths.Event, "Result", widths.Result, "Message", widths.Message)
            : Row("Time", widths.Time, "Type", widths.Type, "Result", widths.Result, "Message", widths.Message);
    }

    private static string FormatConnectivityHistoryRow(GraphTableRow row)
    {
        var widths = GetConnectivityHistoryWidths();
        return widths.Event > 0
            ? Row(
                GetField(row, "eventDateTime"), widths.Time,
                GetField(row, "eventType"), widths.Type,
                GetField(row, "eventName"), widths.Event,
                GetField(row, "eventResult"), widths.Result,
                GetField(row, "message"), widths.Message)
            : Row(
                GetField(row, "eventDateTime"), widths.Time,
                GetField(row, "eventType"), widths.Type,
                GetField(row, "eventResult"), widths.Result,
                GetField(row, "message"), widths.Message);
    }

    private static void RenderConnectivityHistorySummary(IReadOnlyList<GraphTableRow> rows)
    {
        var ordered = rows
            .Select(row => new
            {
                Row = row,
                EventTime = ParseGraphDate(GetField(row, "eventDateTime")),
                EventName = GetField(row, "eventName"),
                EventResult = GetField(row, "eventResult")
            })
            .Where(item => item.EventTime is not null)
            .OrderByDescending(item => item.EventTime)
            .ToArray();

        var latest = ordered.FirstOrDefault();
        var lastStarted = ordered.FirstOrDefault(item =>
            item.EventName.Contains("Started", StringComparison.OrdinalIgnoreCase) &&
            item.EventResult.Contains("success", StringComparison.OrdinalIgnoreCase));
        var lastFinished = ordered.FirstOrDefault(item =>
            item.EventName.Contains("Finished", StringComparison.OrdinalIgnoreCase) &&
            item.EventResult.Contains("success", StringComparison.OrdinalIgnoreCase));
        var inferredState = lastStarted?.EventTime is not null &&
            (lastFinished?.EventTime is null || lastStarted.EventTime > lastFinished.EventTime)
                ? "Possibly connected"
                : lastFinished?.EventTime is not null
                    ? "Last known disconnected"
                    : "Unknown";

        var rowsPanel = new Rows(
            new Markup($"[bold]Latest event[/] {Markup.Escape(FormatConnectivityEvent(latest?.EventName, latest?.EventTime))}"),
            new Markup($"[bold]Last started[/] {Markup.Escape(FormatConnectivityEvent(lastStarted?.EventName, lastStarted?.EventTime))}"),
            new Markup($"[bold]Last finished[/] {Markup.Escape(FormatConnectivityEvent(lastFinished?.EventName, lastFinished?.EventTime))}"),
            new Markup($"[bold]Inferred state[/] {Markup.Escape(inferredState)}"));

        AnsiConsole.Write(new Panel(rowsPanel)
            .Header("Connection summary")
            .Border(BoxBorder.Rounded));
    }

    private static DateTimeOffset? ParseGraphDate(string value)
    {
        return DateTimeOffset.TryParse(value, out var parsed) ? parsed.ToLocalTime() : null;
    }

    private static string FormatConnectivityEvent(string? eventName, DateTimeOffset? eventTime)
    {
        return eventTime is null
            ? "-"
            : $"{eventName ?? "Event"} at {eventTime.Value:g}";
    }

    private static (int Time, int Type, int Event, int Result, int Message) GetConnectivityHistoryWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        var showEvent = available >= 112;
        var time = showEvent ? 22 : 19;
        var type = showEvent ? 16 : 14;
        const int result = 10;
        var eventWidth = showEvent ? 24 : 0;
        var gaps = showEvent ? 4 : 3;
        var message = Math.Max(24, available - time - type - eventWidth - result - gaps);
        return (time, type, eventWidth, result, message);
    }

    private static string GetConnectionHistoryReportHeader()
    {
        var widths = GetConnectionHistoryReportWidths();
        return Row("Session begin", widths.Begin, "Session end", widths.End, "UPN", widths.Upn, "Client OS", widths.ClientOs, "Transport", widths.Transport);
    }

    private static string FormatConnectionHistoryReportRow(GraphTableRow row)
    {
        var widths = GetConnectionHistoryReportWidths();
        return Row(
            GetField(row, "SessionBeginTime"), widths.Begin,
            GetField(row, "SessionEndTime"), widths.End,
            GetField(row, "UPN"), widths.Upn,
            GetField(row, "ClientOS"), widths.ClientOs,
            GetField(row, "TransportType"), widths.Transport);
    }

    private static (int Begin, int End, int Upn, int ClientOs, int Transport) GetConnectionHistoryReportWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int begin = 20;
        const int end = 20;
        const int transport = 12;
        var gaps = 4;
        var remaining = Math.Max(30, available - begin - end - transport - gaps);
        var upn = Math.Max(20, (int)(remaining * 0.6));
        var clientOs = Math.Max(10, remaining - upn);
        return (begin, end, upn, clientOs, transport);
    }

    private static string GetLaunchDetailsHeader()
    {
        var widths = GetLaunchDetailsWidths();
        return Row("Cloud PC", widths.CloudPc, "User", widths.User, "Status", widths.Status, "Switch", widths.Switch);
    }

    private static string FormatLaunchDetailsRow(GraphTableRow row)
    {
        var widths = GetLaunchDetailsWidths();
        return Row(
            GetField(row, "Cloud PC"), widths.CloudPc,
            GetField(row, "User"), widths.User,
            GetField(row, "Status"), widths.Status,
            GetSwitchValue(row), widths.Switch);
    }

    private static (int CloudPc, int User, int Status, int Switch) GetLaunchDetailsWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 12;
        const int switchWidth = 8;
        var remaining = Math.Max(44, available - status - switchWidth - 3);
        var cloudPc = Math.Max(28, (int)(remaining * 0.48));
        var user = Math.Max(18, remaining - cloudPc);
        return (cloudPc, user, status, switchWidth);
    }

    private static string GetServicePlansHeader()
    {
        return Row("Name", 44, "Type", 12, "vCPU", 6, "RAM", 8, "Storage", 10, "Profile", 10);
    }

    private static string FormatServicePlanRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "Name"), 44,
            GetField(row, "Type"), 12,
            GetField(row, "vCPU"), 6,
            GetField(row, "RAM"), 8,
            GetField(row, "Storage"), 10,
            GetField(row, "Profile"), 10);
    }

    private static string GetGalleryImagesHeader()
    {
        var widths = GetCatalogImageWidths();
        return Row("Name", widths.Name, "Status", widths.Status, "Recommended SKU", widths.Sku, "Size", widths.Size, "OS version", widths.Os);
    }

    private static string FormatGalleryImageRow(GraphTableRow row)
    {
        var widths = GetCatalogImageWidths();
        return Row(
            GetField(row, "displayName"), widths.Name,
            GetField(row, "status"), widths.Status,
            GetField(row, "recommendedSku"), widths.Sku,
            FormatCatalogGb(GetField(row, "sizeInGB")), widths.Size,
            GetField(row, "osVersionNumber"), widths.Os);
    }

    private static string GetCustomImagesHeader()
    {
        var widths = GetCatalogImageWidths();
        return Row("Name", widths.Name, "Status", widths.Status, "OS", widths.Sku, "Size", widths.Size, "Modified", widths.Os);
    }

    private static string FormatCustomImageRow(GraphTableRow row)
    {
        var widths = GetCatalogImageWidths();
        return Row(
            GetField(row, "displayName"), widths.Name,
            GetField(row, "status"), widths.Status,
            GetField(row, "operatingSystem"), widths.Sku,
            FormatCatalogGb(GetField(row, "sizeInGB")), widths.Size,
            GetField(row, "lastModifiedDateTime"), widths.Os);
    }

    private static (int Name, int Status, int Sku, int Size, int Os) GetCatalogImageWidths()
    {
        var available = Math.Max(92, Console.WindowWidth - 4);
        const int status = 14;
        const int size = 8;
        var remaining = Math.Max(50, available - status - size - 4);
        var name = Math.Clamp((int)(remaining * 0.42), 30, 44);
        var sku = Math.Clamp((int)(remaining * 0.34), 18, 30);
        var os = Math.Max(16, remaining - name - sku);
        return (name, status, sku, size, os);
    }

    private static string GetSupportedRegionsHeader()
    {
        return Row("Name", 34, "Status", 12, "Solution", 16, "Group", 20, "Geo", 20);
    }

    private static string FormatSupportedRegionRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 34,
            GetField(row, "regionStatus"), 12,
            GetField(row, "supportedSolution"), 16,
            GetField(row, "regionGroup"), 20,
            GetField(row, "geographicLocationType"), 20);
    }

    private static string FormatCatalogGb(string value)
    {
        return value == "-" ? "-" : $"{value} GB";
    }

    private static string GetOrganizationSettingsHeader()
    {
        return Row("OS", 14, "User", 18, "MEM auto", 10, "SSO", 8, "Language", 14);
    }

    private static string FormatOrganizationSettingRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "osVersion"), 14,
            GetField(row, "userAccountType"), 18,
            FormatBooleanCell(GetField(row, "memAutoEnrollEnabled")), 10,
            FormatBooleanCell(GetField(row, "singleSignOnEnabled")), 8,
            GetField(row, "windowsLanguage"), 14);
    }

    private static string GetSettingProfilesHeader()
    {
        return Row("Name", 42, "Type", 18, "Assigned", 10, "Priority", 10, "Modified", 20);
    }

    private static string FormatSettingProfileRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 42,
            GetField(row, "profileType"), 18,
            FormatBooleanCell(GetField(row, "isAssigned")), 10,
            GetNestedField(row, "priorityMetaData", "priority"), 10,
            FormatDateCell(GetField(row, "lastModifiedDateTime")), 20);
    }

    private static string GetUserSettingsHeader()
    {
        return Row("Name", 38, "Self svc", 10, "Admin", 8, "Reset", 8, "Restore", 9, "DR", 8);
    }

    private static string FormatUserSettingRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 38,
            FormatBooleanCell(GetField(row, "selfServiceEnabled")), 10,
            FormatBooleanCell(GetField(row, "localAdminEnabled")), 8,
            FormatBooleanCell(GetField(row, "resetEnabled")), 8,
            FormatBooleanCell(GetNestedField(row, "restorePointSetting", "userRestoreEnabled")), 9,
            FormatBooleanCell(GetNestedField(row, "crossRegionDisasterRecoverySetting", "crossRegionDisasterRecoveryEnabled")), 8);
    }

    private static string GetNestedField(GraphTableRow row, string objectName, string propertyName)
    {
        var value = GetOptionalField(row, propertyName);
        if (!string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        var raw = GetOptionalField(row, objectName);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return "-";
        }

        try
        {
            using var document = System.Text.Json.JsonDocument.Parse(raw);
            if (document.RootElement.TryGetProperty(propertyName, out var property))
            {
                return property.ValueKind switch
                {
                    System.Text.Json.JsonValueKind.String => property.GetString() ?? "-",
                    System.Text.Json.JsonValueKind.True => "true",
                    System.Text.Json.JsonValueKind.False => "false",
                    System.Text.Json.JsonValueKind.Number => property.GetRawText(),
                    _ => "-"
                };
            }
        }
        catch (System.Text.Json.JsonException)
        {
            return "-";
        }

        return "-";
    }

    private static string FormatBooleanCell(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "true" => "Yes",
            "false" => "No",
            _ => value
        };
    }

    private static string FormatDateCell(string value)
    {
        if (DateTimeOffset.TryParse(value, out var parsed))
        {
            return parsed.ToLocalTime().ToString("g");
        }

        return value;
    }

    private static string GetSwitchValue(GraphTableRow row)
    {
        var value = GetOptionalField(row, "windows365SwitchCompatible", "Windows365SwitchCompatible");
        return value?.ToLowerInvariant() switch
        {
            "true" => "Yes",
            "false" => "No",
            null => "-",
            _ => value
        };
    }

    private static string GetField(GraphTableRow row, string name)
    {
        return row.Fields.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : "-";
    }

    private static string? GetOptionalField(GraphTableRow row, params string[] names)
    {
        foreach (var name in names)
        {
            if (row.Fields.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value) && value != "-")
            {
                return value;
            }
        }

        return null;
    }

    private static void ShowGraphRowDetails(string title, GraphTableRow row)
    {
        AnsiConsole.Clear();
        var lines = row.Fields
            .OrderBy(field => field.Key, StringComparer.OrdinalIgnoreCase)
            .Select(field => new Markup(PropertyBlock(field.Key, field.Value)));

        var panel = new Panel(new Rows(lines))
            .Header(title)
            .Border(BoxBorder.Rounded);

        AnsiConsole.Write(panel);
        WaitForBack();
    }

    private async Task ShowCloudPcsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var cloudPcs = await LoadCloudPcsAsync();

        if (cloudPcs.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No Cloud PCs returned.[/]");
            Pause();
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = CloudPcSortMode.Name;

        while (true)
        {
            var visibleCloudPcs = SortCloudPcs(FilterCloudPcs(cloudPcs, filter), sortMode);
            if (selectedIndex >= visibleCloudPcs.Count)
            {
                selectedIndex = Math.Max(0, visibleCloudPcs.Count - 1);
            }

            RenderCloudPcBrowser(cloudPcs, visibleCloudPcs, selectedIndex, filter, sortMode);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleCloudPcs.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visibleCloudPcs.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visibleCloudPcs.Count - 1);
                    break;
                case ConsoleKey.Enter:
                    if (visibleCloudPcs.Count > 0)
                    {
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex]);
                    }
                    break;
                case ConsoleKey.A:
                    if (visibleCloudPcs.Count > 0)
                    {
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex]);
                    }
                    break;
                case ConsoleKey.R:
                    cloudPcs = await LoadCloudPcsAsync();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextCloudPcSortMode(sortMode);
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (IsActionHistoryHotkey(key))
                    {
                        await ShowActionHistoryAsync();
                    }
                    else if (key.KeyChar == '/' || key.KeyChar == 'f' || key.KeyChar == 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar == 'q' || key.KeyChar == 'Q' || key.KeyChar == 'b' || key.KeyChar == 'B')
                    {
                        return;
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'd' or 'D')
                    {
                        await ShowDiskSpaceAsync(visibleCloudPcs[selectedIndex]);
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'n' or 'N')
                    {
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex], "Snapshots");
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'z' or 'Z')
                    {
                        await ShowResizeAsync(visibleCloudPcs[selectedIndex]);
                    }
                    else if (visibleCloudPcs.Count > 0 && key.KeyChar is 'y' or 'Y')
                    {
                        await InvokeCloudPcActionAsync(visibleCloudPcs[selectedIndex], "Sync");
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private async Task ShowCloudAppsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        IReadOnlyList<CloudAppSummary> apps;
        try
        {
            apps = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading Cloud Apps...", async _ => await _session.Graph.GetCloudAppsAsync());
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine("[red]Failed to load Cloud Apps.[/]");
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            Pause();
            return;
        }

        if (apps.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No Cloud Apps returned.[/]");
            Pause();
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;

        while (true)
        {
            var visibleApps = FilterCloudApps(apps, filter);
            if (selectedIndex >= visibleApps.Count)
            {
                selectedIndex = Math.Max(0, visibleApps.Count - 1);
            }

            RenderCloudAppBrowser(apps, visibleApps, selectedIndex, filter);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleApps.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visibleApps.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visibleApps.Count - 1);
                    break;
                case ConsoleKey.Enter:
                case ConsoleKey.A:
                    if (visibleApps.Count > 0)
                    {
                        await ShowCloudAppDetailsAsync(visibleApps[selectedIndex]);
                    }
                    break;
                case ConsoleKey.R:
                    apps = await LoadCloudAppsAsync();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (key.KeyChar == '/' || key.KeyChar == 'f' || key.KeyChar == 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar == 'q' || key.KeyChar == 'Q' || key.KeyChar == 'b' || key.KeyChar == 'B')
                    {
                        return;
                    }
                    else if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private static T? SelectFromTable<T>(
        string title,
        string header,
        IReadOnlyList<T> items,
        Func<T, string> rowFactory)
    {
        var rows = items
            .Select(item => new TableChoice<T>(rowFactory(item), item, false))
            .Concat([new TableChoice<T>("Back", default, true)])
            .ToArray();

        var selected = AnsiConsole.Prompt(
            new SelectionPrompt<TableChoice<T>>()
                .Title($"[#58a6ff]{Markup.Escape(title)}[/]\n[grey]{Markup.Escape(header)}[/]\n[grey]{Markup.Escape(new string('-', header.Length))}[/]")
                .HighlightStyle(SelectionHighlightStyle())
                .PageSize(18)
                .UseConverter(choice => Markup.Escape(choice.Label))
                .AddChoices(rows));

        return selected.IsBack ? default : selected.Item;
    }

    private async Task<IReadOnlyList<CloudPcSummary>> LoadCloudPcsAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud PCs...", async _ => await _session.Graph.GetCloudPcsAsync());
    }

    private async Task<IReadOnlyList<CloudAppSummary>> LoadCloudAppsAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud Apps...", async _ => await _session.Graph.GetCloudAppsAsync());
    }

    private void RenderCloudPcBrowser(
        IReadOnlyList<CloudPcSummary> allCloudPcs,
        IReadOnlyList<CloudPcSummary> visibleCloudPcs,
        int selectedIndex,
        string filter,
        CloudPcSortMode sortMode)
    {
        AnsiConsole.Clear();

        var selectedCloudPc = visibleCloudPcs.Count > 0 ? visibleCloudPcs[selectedIndex] : null;
        var grid = new Grid();
        grid.AddColumn();
        grid.AddColumn();
        grid.AddRow(CreateCloudPcTable(allCloudPcs, visibleCloudPcs, selectedIndex, filter), CreateCloudPcSidePanel(selectedCloudPc));

        RenderBreadcrumb("Cloud PCs", "Browse");
        AnsiConsole.Write(CreateCloudPcSummaryPanel(allCloudPcs, visibleCloudPcs, filter));
        if (Console.WindowWidth >= 125)
        {
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(CreateCloudPcTable(allCloudPcs, visibleCloudPcs, selectedIndex, filter));
            AnsiConsole.Write(CreateCloudPcSidePanel(selectedCloudPc));
        }
        AnsiConsole.MarkupLine($"[grey]Sort: {FormatCloudPcSortMode(sortMode)} | Up/Down move | PgUp/PgDn page | Enter actions | D disk | N snapshots | Z resize | Y sync | / filter | C clear | S sort | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    private static Panel CreateCloudPcSummaryPanel(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, string filter)
    {
        var statusSummary = string.Join("  ", allCloudPcs
            .GroupBy(pc => pc.Status ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));

        var typeSummary = string.Join("  ", allCloudPcs
            .GroupBy(pc => pc.ProvisioningType ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));

        var rows = new List<Spectre.Console.Rendering.IRenderable>
        {
            new Markup($"[bold]Total[/] {allCloudPcs.Count}   [bold]Visible[/] {visibleCloudPcs.Count}   [bold]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
            new Markup($"[bold]Status[/] {Markup.Escape(statusSummary)}"),
            new Markup($"[bold]Type[/] {Markup.Escape(typeSummary)}")
        };

        if (allCloudPcs.Any(pc => pc.ConnectivityResult is not null))
        {
            var inUseCount = allCloudPcs.Count(pc => string.Equals(pc.ConnectivityResult?.Status, "inUse", StringComparison.OrdinalIgnoreCase));
            var availableCount = allCloudPcs.Count(pc => string.Equals(pc.ConnectivityResult?.Status, "available", StringComparison.OrdinalIgnoreCase));
            rows.Add(new Markup($"[bold]Shared usage[/] [yellow]In use: {inUseCount}[/]  [green]Available: {availableCount}[/]"));
        }

        return new Panel(new Rows(rows)).Border(BoxBorder.Rounded).Header("Cloud PC fleet");
    }

    private static Table CreateCloudPcTable(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, int selectedIndex, string filter)
    {
        var widths = GetCloudPcWidths();
        var showInUse = allCloudPcs.Any(pc => pc.ConnectivityResult is not null);
        var table = new Table()
            .Title("Cloud PCs")
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Status")
            .AddColumn("Type")
            .AddColumn("Name");

        var showUser = Console.WindowWidth >= 105;
        var showServicePlan = Console.WindowWidth >= 135;

        if (showInUse)
        {
            table.AddColumn("In use");
        }

        if (showUser)
        {
            table.AddColumn("User");
        }

        if (showServicePlan)
        {
            table.AddColumn("Service plan");
        }

        if (visibleCloudPcs.Count == 0)
        {
            var emptyCells = new List<string> { "-", "-", "-", "[grey]No Cloud PCs match the current filter.[/]" };
            if (showInUse)
            {
                emptyCells.Add("-");
            }
            if (showUser)
            {
                emptyCells.Add("-");
            }
            if (showServicePlan)
            {
                emptyCells.Add("-");
            }
            table.AddRow(emptyCells.ToArray());
            return table;
        }

        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 15));
        var start = Math.Max(0, Math.Min(selectedIndex - pageSize / 2, Math.Max(0, visibleCloudPcs.Count - pageSize)));
        var end = Math.Min(visibleCloudPcs.Count - 1, start + pageSize - 1);

        for (var index = start; index <= end; index++)
        {
            var pc = visibleCloudPcs[index];
            var selected = index == selectedIndex;
            var row = new List<string>
            {
                selected ? "[black on #58a6ff]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(pc.Status ?? "unknown", widths.Status))) : StatusMarkup(pc.Status, widths.Status),
                selected ? Selected(Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type))) : Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type)),
                selected ? Selected(Markup.Escape(Fit(pc.Name, widths.Name))) : Markup.Escape(Fit(pc.Name, widths.Name))
            };

            if (showInUse)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(FormatInUsePlain(pc), 26))) : FormatInUseMarkup(pc));
            }

            if (showUser)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(pc.EffectiveUserPrincipalName ?? "-", widths.User))) : Markup.Escape(Fit(pc.EffectiveUserPrincipalName ?? "-", widths.User)));
            }

            if (showServicePlan)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(pc.ServicePlanName ?? "-", widths.ServicePlan))) : Markup.Escape(Fit(pc.ServicePlanName ?? "-", widths.ServicePlan)));
            }

            table.AddRow(row.ToArray());
        }

        return table;
    }

    private static string FormatInUsePlain(CloudPcSummary pc)
    {
        var status = pc.ConnectivityResult?.Status;
        if (string.IsNullOrWhiteSpace(status))
        {
            return "-";
        }

        if (string.Equals(status, "inUse", StringComparison.OrdinalIgnoreCase))
        {
            var sessionStart = pc.SharedDeviceDetail?.SessionStartDateTime;
            return sessionStart is null
                ? "In use"
                : $"In use since {sessionStart.Value.ToLocalTime():t}";
        }

        if (string.Equals(status, "available", StringComparison.OrdinalIgnoreCase))
        {
            return "Available";
        }

        return status;
    }

    private static string FormatInUseMarkup(CloudPcSummary pc)
    {
        var status = pc.ConnectivityResult?.Status;
        var text = Markup.Escape(Fit(FormatInUsePlain(pc), 26));

        if (string.Equals(status, "inUse", StringComparison.OrdinalIgnoreCase))
        {
            return $"[yellow]{text}[/]";
        }

        if (string.Equals(status, "available", StringComparison.OrdinalIgnoreCase))
        {
            return $"[green]{text}[/]";
        }

        return $"[grey]{text}[/]";
    }

    private static Panel CreateCloudPcSidePanel(CloudPcSummary? cloudPc)
    {
        if (cloudPc is null)
        {
            return new Panel("[grey]No Cloud PC selected.[/]")
                .Header("Details")
                .Border(BoxBorder.Rounded);
        }

        var content = new Rows(
            new Markup(PropertyBlock("Name", cloudPc.Name, "grey")),
            new Markup(PropertyInline("Status", StatusMarkup(cloudPc.Status), valueIsMarkup: true)),
            new Markup(PropertyInline("Type", cloudPc.ProvisioningType ?? "-", "grey")),
            new Markup(PropertyBlock("User", cloudPc.EffectiveUserPrincipalName ?? "-", "grey")),
            new Markup(PropertyBlock("Service plan", cloudPc.ServicePlanName ?? "-", "grey")),
            new Markup(PropertyInline("In use", FormatInUseMarkup(cloudPc), valueIsMarkup: true)),
            new Markup(PropertyBlock("Cloud PC ID", cloudPc.Id, "grey")),
            new Markup(PropertyBlock("Actions", "Enter details, A actions", "grey")));

        return new Panel(content)
            .Header("Selected Cloud PC")
            .Border(BoxBorder.Rounded);
    }

    private void RenderCloudAppBrowser(
        IReadOnlyList<CloudAppSummary> allApps,
        IReadOnlyList<CloudAppSummary> visibleApps,
        int selectedIndex,
        string filter)
    {
        AnsiConsole.Clear();
        var selectedApp = visibleApps.Count > 0 ? visibleApps[selectedIndex] : null;
        var grid = new Grid();
        grid.AddColumn();
        grid.AddColumn();
        grid.AddRow(CreateCloudAppTable(allApps, visibleApps, selectedIndex, filter), CreateCloudAppSidePanel(selectedApp));

        RenderBreadcrumb("Cloud Apps", "Browse");
        AnsiConsole.Write(CreateCloudAppSummaryPanel(allApps, visibleApps, filter));
        if (Console.WindowWidth >= 125)
        {
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(CreateCloudAppTable(allApps, visibleApps, selectedIndex, filter));
            AnsiConsole.Write(CreateCloudAppSidePanel(selectedApp));
        }
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter details | A actions | / filter | C clear | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    private static Panel CreateCloudAppSummaryPanel(IReadOnlyList<CloudAppSummary> allApps, IReadOnlyList<CloudAppSummary> visibleApps, string filter)
    {
        var statusSummary = string.Join("  ", allApps
            .GroupBy(app => app.AppStatus ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));

        var content = new Rows(
            new Markup($"[bold]Total[/] {allApps.Count}   [bold]Visible[/] {visibleApps.Count}   [bold]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
            new Markup($"[bold]Status[/] {Markup.Escape(statusSummary)}"));

        return new Panel(content).Border(BoxBorder.Rounded).Header("Cloud Apps");
    }

    private static Table CreateCloudAppTable(IReadOnlyList<CloudAppSummary> allApps, IReadOnlyList<CloudAppSummary> visibleApps, int selectedIndex, string filter)
    {
        var widths = GetCloudAppWidths();
        var table = new Table()
            .Title("Cloud Apps")
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Status")
            .AddColumn("Name");

        var showPublisher = Console.WindowWidth >= 100;
        var showDates = Console.WindowWidth >= 150;
        if (showPublisher)
        {
            table.AddColumn("Publisher");
        }
        if (showDates)
        {
            table.AddColumn("Published");
            table.AddColumn("Added");
        }

        if (visibleApps.Count == 0)
        {
            var emptyCells = new List<string> { "-", "-", "[grey]No Cloud Apps match the current filter.[/]" };
            if (showPublisher) { emptyCells.Add("-"); }
            if (showDates) { emptyCells.Add("-"); emptyCells.Add("-"); }
            table.AddRow(emptyCells.ToArray());
            return table;
        }

        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 15));
        var start = Math.Max(0, Math.Min(selectedIndex - pageSize / 2, Math.Max(0, visibleApps.Count - pageSize)));
        var end = Math.Min(visibleApps.Count - 1, start + pageSize - 1);

        for (var index = start; index <= end; index++)
        {
            var app = visibleApps[index];
            var selected = index == selectedIndex;
            var row = new List<string>
            {
                selected ? "[black on #58a6ff]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(app.AppStatus ?? "unknown", widths.Status))) : AppStatusMarkup(app.AppStatus, widths.Status),
                selected ? Selected(Markup.Escape(Fit(app.DisplayName, widths.Name))) : Markup.Escape(Fit(app.DisplayName, widths.Name))
            };
            if (showPublisher)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(app.Publisher ?? "-", widths.Publisher))) : Markup.Escape(Fit(app.Publisher ?? "-", widths.Publisher)));
            }
            if (showDates)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Published))) : Markup.Escape(Fit(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Published)));
                row.Add(selected ? Selected(Markup.Escape(Fit(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Added))) : Markup.Escape(Fit(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-", widths.Added)));
            }
            table.AddRow(row.ToArray());
        }

        return table;
    }

    private static Panel CreateCloudAppSidePanel(CloudAppSummary? app)
    {
        if (app is null)
        {
            return new Panel("[grey]No Cloud App selected.[/]").Header("Details").Border(BoxBorder.Rounded);
        }

        var content = new Rows(
            new Markup($"[bold]Name[/]\n{Markup.Escape(app.DisplayName)}"),
            new Markup($"[bold]Status[/] {AppStatusMarkup(app.AppStatus, 12)}"),
            new Markup($"[bold]Publisher[/]\n{Markup.Escape(app.Publisher ?? "-")}"),
            new Markup($"[bold]Added[/] {Markup.Escape(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
            new Markup($"[bold]Published[/] {Markup.Escape(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
            new Markup($"[bold]Cloud app ID[/]\n[grey]{Markup.Escape(app.Id)}[/]"),
            new Markup("[bold]Actions[/]\n[grey]Press A to open actions for this Cloud App.[/]"));

        return new Panel(content).Header("Selected Cloud App").Border(BoxBorder.Rounded);
    }

    private static IReadOnlyList<CloudAppSummary> FilterCloudApps(IReadOnlyList<CloudAppSummary> apps, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return apps;
        }

        return apps
            .Where(app =>
                Contains(app.DisplayName, filter) ||
                Contains(app.AppStatus, filter) ||
                Contains(app.Publisher, filter) ||
                Contains(app.DiscoveredAppName, filter))
            .ToArray();
    }

    private static IReadOnlyList<CloudPcSummary> FilterCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return cloudPcs;
        }

        return cloudPcs
            .Where(pc =>
                Contains(pc.Name, filter) ||
                Contains(pc.Status, filter) ||
                Contains(pc.ProvisioningType, filter) ||
                Contains(pc.EffectiveUserPrincipalName, filter) ||
                Contains(pc.ServicePlanName, filter))
            .ToArray();
    }

    private static IReadOnlyList<CloudPcSummary> SortCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Status => cloudPcs.OrderBy(pc => pc.Status, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            CloudPcSortMode.User => cloudPcs.OrderBy(pc => pc.EffectiveUserPrincipalName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            CloudPcSortMode.ServicePlan => cloudPcs.OrderBy(pc => pc.ServicePlanName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => cloudPcs.OrderBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray()
        };
    }

    private static CloudPcSortMode NextCloudPcSortMode(CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Name => CloudPcSortMode.Status,
            CloudPcSortMode.Status => CloudPcSortMode.User,
            CloudPcSortMode.User => CloudPcSortMode.ServicePlan,
            _ => CloudPcSortMode.Name
        };
    }

    private static string FormatCloudPcSortMode(CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Status => "status",
            CloudPcSortMode.User => "user",
            CloudPcSortMode.ServicePlan => "service plan",
            _ => "name"
        };
    }

    private static IReadOnlyList<ProvisioningPolicySummary> FilterProvisioningPolicies(IReadOnlyList<ProvisioningPolicySummary> policies, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return policies;
        }

        return policies
            .Where(policy =>
                Contains(policy.DisplayName, filter) ||
                Contains(policy.Description, filter) ||
                Contains(policy.ProvisioningType, filter) ||
                Contains(policy.ImageDisplayName, filter) ||
                Contains(policy.ImageType, filter) ||
                Contains(policy.DomainJoinTypes, filter) ||
                Contains(policy.CloudPcNamingTemplate, filter) ||
                Contains(policy.CloudPcGroupDisplayName, filter) ||
                policy.AssignedGroupNames.Any(groupName => Contains(groupName, filter)))
            .ToArray();
    }

    private static IReadOnlyList<ProvisioningPolicySummary> SortProvisioningPolicies(IReadOnlyList<ProvisioningPolicySummary> policies, ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Type => policies.OrderBy(policy => policy.ProvisioningType, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            ProvisioningPolicySortMode.Image => policies.OrderBy(policy => policy.ImageDisplayName, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            ProvisioningPolicySortMode.Join => policies.OrderBy(policy => policy.DomainJoinTypes, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => policies.OrderBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray()
        };
    }

    private static ProvisioningPolicySortMode NextProvisioningPolicySortMode(ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Name => ProvisioningPolicySortMode.Type,
            ProvisioningPolicySortMode.Type => ProvisioningPolicySortMode.Image,
            ProvisioningPolicySortMode.Image => ProvisioningPolicySortMode.Join,
            _ => ProvisioningPolicySortMode.Name
        };
    }

    private static string FormatProvisioningPolicySortMode(ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Type => "type",
            ProvisioningPolicySortMode.Image => "image",
            ProvisioningPolicySortMode.Join => "join",
            _ => "name"
        };
    }

    private static bool Contains(string? value, string filter)
    {
        return value?.Contains(filter, StringComparison.OrdinalIgnoreCase) == true;
    }

    private static string[] GetCloudPcActions(CloudPcSummary cloudPc)
    {
        var actions = new List<string>();

        if (HasStatusDetailsWorthViewing(cloudPc))
        {
            actions.Add("View status details");
        }

        actions.AddRange([
            "Remote action history",
            "Connection history",
            "Disk space",
            "Snapshots",
            "Resize"
        ]);

        if (IsCloudPcClearlyOff(cloudPc))
        {
            actions.Add("Power on");
        }

        actions.AddRange([
            "Rename",
            "Sync",
            "Restart",
            "Reset local admin password"
        ]);

        if (IsCloudPcInGracePeriod(cloudPc))
        {
            actions.Add("End grace period");
        }

        actions.AddRange([
            "Reprovision",
            "Back"
        ]);

        return actions.ToArray();
    }

    /// <summary>
    /// "View status details" is only useful when the Cloud PC's status indicates something didn't
    /// apply cleanly — mirrors the Intune portal, which only shows "View more information" for
    /// these states.
    /// </summary>
    private static bool HasStatusDetailsWorthViewing(CloudPcSummary cloudPc)
    {
        return MatchesAny(cloudPc.Status, "provisionedwithwarnings", "provisionedwitherrors", "failed");
    }

    private static bool IsCloudPcClearlyOff(CloudPcSummary cloudPc)
    {
        return MatchesAny(cloudPc.PowerState, "off", "stopped", "deallocated", "poweredoff") ||
            MatchesAny(cloudPc.Status, "off", "stopped", "deallocated", "poweredoff");
    }

    private static bool IsCloudPcInGracePeriod(CloudPcSummary cloudPc)
    {
        return string.Equals(cloudPc.Status, "inGracePeriod", StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesAny(string? value, params string[] terms)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return terms.Any(term => value.Contains(term, StringComparison.OrdinalIgnoreCase));
    }

    private static string PromptFilter()
    {
        AnsiConsole.WriteLine();
        return AnsiConsole.Ask<string>("Filter:");
    }

    private static string StatusMarkup(string? status, int width = 24)
    {
        var text = status ?? "unknown";
        var color = text.ToLowerInvariant() switch
        {
            "provisioned" or "available" or "ready" => "darkolivegreen3_1",
            "provisionedwithwarnings" or "provisionedwitherrors" => "orange1",
            "provisioning" or "pending" or "inprogress" or "pendingprovision" or "pendingupdate" => "khaki1",
            "failed" or "error" or "resizevalidationfailed" or "movingregionfailed" or "restorefailed" => "indianred1",
            "ingraceperiod" => "plum1",
            _ => "grey"
        };

        return $"[{color}]{Markup.Escape(Fit(text, width))}[/]";
    }

    private static string AppStatusMarkup(string? status, int width)
    {
        var text = status ?? "unknown";
        var color = text.ToLowerInvariant() switch
        {
            "ready" => "khaki1",
            "published" => "darkolivegreen3_1",
            "failed" => "indianred1",
            _ => "grey"
        };

        return $"[{color}]{Markup.Escape(Fit(text, width))}[/]";
    }

    private static string PropertyInline(string name, string value, string valueColor = "grey", bool valueIsMarkup = false)
    {
        var renderedValue = valueIsMarkup ? value : $"[{valueColor}]{Markup.Escape(value)}[/]";
        return $"[white]{Markup.Escape(name)}:[/] {renderedValue}";
    }

    private static string PropertyBlock(string name, string value, string valueColor = "grey")
    {
        return $"[white]{Markup.Escape(name)}[/]\n[{valueColor}]{Markup.Escape(value)}[/]";
    }

    private static string Selected(string escapedText)
    {
        return $"[black on {AccentColor}]{escapedText}[/]";
    }

    private static void SetStatus(string markup)
    {
        statusMessage = markup;
        statusMessageAt = DateTimeOffset.Now;
    }

    private void UpdateStatusBarSnapshot()
    {
        var connectionText = _session.IsConnected
            ? "[black on #3fb950] CONNECTED [/]"
            : "[white on red] NOT CONNECTED [/]";
        var tenantName = _session.TenantName ?? "No tenant selected";
        var tenantId = _session.TenantId ?? "-";
        statusBarConnection = connectionText;
        statusBarTenant = tenantId == "-"
            ? tenantName
            : $"{tenantName} ({tenantId})";
    }

    private static void RenderStatusBar()
    {
        RenderHomeStatusLine();
    }

    private static bool IsActionHistoryHotkey(ConsoleKeyInfo key)
    {
        return key.KeyChar is 'h' or 'H';
    }

    private static void AddActionHistory(string action, string target, string status, string? detail = null, string resourceType = "Cloud PC", string? resourceName = null)
    {
        ActionHistory.Insert(0, new ActionHistoryItem(action, target, resourceType, resourceName ?? InferResourceName(target), status, DateTimeOffset.Now, detail));
        if (ActionHistory.Count > 100)
        {
            ActionHistory.RemoveRange(100, ActionHistory.Count - 100);
        }
    }

    private static string? InferResourceName(string target)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            return null;
        }

        var markerIndex = target.IndexOf(" to ", StringComparison.OrdinalIgnoreCase);
        return markerIndex > 0 ? target[..markerIndex] : target;
    }

    private static Style SelectionHighlightStyle()
    {
        return new Style(Color.Black, Color.FromHex(AccentColor));
    }

    /// <summary>
    /// Standard yes/no confirmation UI for the whole app — an arrow-key selection prompt rather
    /// than a type-y/n-and-Enter prompt, so every confirmation in the CLI behaves identically.
    /// </summary>
    private static bool AskYesNo(string question, bool defaultToYes = true)
    {
        var choices = defaultToYes ? new[] { "Yes", "No" } : new[] { "No", "Yes" };
        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title(question)
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices(choices));
        return choice == "Yes";
    }

    private static readonly string[] DestructiveActionNames = ["Delete"];

    /// <summary>
    /// Renders one line of an action list, highlighting destructive actions (e.g. "Delete") in
    /// red so they visually stand out from safe actions before the user even reaches a confirm
    /// prompt — same treatment whether or not the row is currently selected.
    /// </summary>
    private static string FormatActionLine(string action, bool selected)
    {
        var isDestructive = DestructiveActionNames.Contains(action, StringComparer.OrdinalIgnoreCase);
        var escaped = Markup.Escape(action);
        if (selected)
        {
            return isDestructive
                ? $"[black on {AccentColor}]> [red]{escaped}[/][/]"
                : $"[black on {AccentColor}]> {escaped}[/]";
        }

        return isDestructive ? $"  [red]{escaped}[/]" : $"  {escaped}";
    }

    private static string Row(params object[] valuesAndWidths)
    {
        if (valuesAndWidths.Length % 2 != 0)
        {
            throw new ArgumentException("Rows require value and width pairs.", nameof(valuesAndWidths));
        }

        var cells = new List<string>();
        for (var index = 0; index < valuesAndWidths.Length; index += 2)
        {
            var value = valuesAndWidths[index]?.ToString() ?? "-";
            var width = Convert.ToInt32(valuesAndWidths[index + 1]);
            cells.Add(Fit(value, width));
        }

        return string.Join(" ", cells);
    }

    private static string Fit(string value, int width)
    {
        if (value.Length > width)
        {
            return string.Concat(value.AsSpan(0, Math.Max(0, width - 3)), "...");
        }

        return value.PadRight(width);
    }

    private static (int Name, int Status, int Type, int User, int ServicePlan) GetCloudPcWidths()
    {
        var available = Math.Max(90, Console.WindowWidth - 4);
        // Wide enough for the longest real cloudPcStatus value ("provisionedWithWarnings" /
        // "resizeValidationFailed" = 23 chars) so status text is never truncated.
        const int status = 24;
        const int type = 10;
        var remaining = Math.Max(40, available - status - type - 4);
        var name = Console.WindowWidth < 105 ? Math.Max(28, remaining - 4) : Math.Max(24, (int)(remaining * 0.32));
        var user = Math.Max(24, (int)(remaining * 0.34));
        var servicePlan = Math.Max(24, remaining - name - user);
        return (name, status, type, user, servicePlan);
    }

    private static (int Status, int Name, int Publisher, int Published, int Added) GetCloudAppWidths()
    {
        const int status = 12;
        const int published = 18;
        const int added = 18;
        var showDates = Console.WindowWidth >= 150;
        var reserved = status + (showDates ? published + added : 0);
        var available = Math.Min(Math.Max(72, Console.WindowWidth - 4), showDates ? 132 : 104);
        var remaining = Math.Max(40, available - reserved - (showDates ? 4 : 2));
        var name = Math.Clamp((int)(remaining * 0.58), 30, 48);
        var publisher = Math.Clamp(remaining - name, 18, 34);
        return (status, name, publisher, published, added);
    }

    private async Task ShowCloudPcDetailsAsync(CloudPcSummary cloudPc, string initialSubPanel = "Actions")
    {
        var actions = GetCloudPcActions(cloudPc);
        var selectedActionIndex = 0;
        CloudPcDiskSpace? diskSpace = null;
        GraphTableRow? signInStatus = await LoadSignInStatusForCloudPcAsync(cloudPc);
        GraphTableRow? latestSession = await LoadLatestSessionForCloudPcAsync(cloudPc);
        IReadOnlyList<CloudPcSnapshot>? snapshots = null;
        IReadOnlyList<CloudPcRemoteActionResult>? remoteActions = null;
        var selectedSnapshotIndex = 0;
        var selectedRemoteActionIndex = 0;
        var activeSubPanel = initialSubPanel;
        if (activeSubPanel == "Remote action history")
        {
            remoteActions = await LoadRemoteActionsForCloudPcAsync(cloudPc);
        }
        else if (activeSubPanel == "Snapshots")
        {
            snapshots = await LoadSnapshotsForCloudPcAsync(cloudPc);
        }
        else if (activeSubPanel == "Disk space")
        {
            diskSpace = await LoadDiskSpaceForCloudPcAsync(cloudPc);
        }

        while (true)
        {
            AnsiConsole.Clear();
            RenderCloudPcDetailLayout(cloudPc, actions, selectedActionIndex, activeSubPanel, diskSpace, signInStatus, latestSession, snapshots, selectedSnapshotIndex, remoteActions, selectedRemoteActionIndex);
            var key = ReadNavigationKey(intercept: true);

            if (activeSubPanel == "Snapshots")
            {
                switch (key.Key)
                {
                    case ConsoleKey.UpArrow:
                        selectedSnapshotIndex = Math.Max(0, selectedSnapshotIndex - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        selectedSnapshotIndex = Math.Min(Math.Max(0, (snapshots?.Count ?? 0) - 1), selectedSnapshotIndex + 1);
                        break;
                    case ConsoleKey.C:
                    case ConsoleKey.N:
                        snapshots = await CreateSnapshotAndReloadAsync(cloudPc);
                        selectedSnapshotIndex = 0;
                        break;
                    case ConsoleKey.R:
                        snapshots = await LoadSnapshotsForCloudPcAsync(cloudPc);
                        selectedSnapshotIndex = 0;
                        break;
                    case ConsoleKey.Enter:
                        if (snapshots is { Count: > 0 })
                        {
                            await ShowSnapshotActionMenuAsync(cloudPc, snapshots[selectedSnapshotIndex]);
                            snapshots = await LoadSnapshotsForCloudPcAsync(cloudPc);
                        }
                        break;
                    case ConsoleKey.Escape:
                    case ConsoleKey.LeftArrow:
                        activeSubPanel = "Actions";
                        break;
                    default:
                        if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                        {
                            activeSubPanel = "Actions";
                        }
                        break;
                }

                continue;
            }

            if (activeSubPanel == "Remote action history")
            {
                switch (key.Key)
                {
                    case ConsoleKey.UpArrow:
                        selectedRemoteActionIndex = Math.Max(0, selectedRemoteActionIndex - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        selectedRemoteActionIndex = Math.Min(Math.Max(0, (remoteActions?.Count ?? 0) - 1), selectedRemoteActionIndex + 1);
                        break;
                    case ConsoleKey.R:
                        remoteActions = await LoadRemoteActionsForCloudPcAsync(cloudPc);
                        selectedRemoteActionIndex = 0;
                        break;
                    case ConsoleKey.Escape:
                    case ConsoleKey.LeftArrow:
                        activeSubPanel = "Actions";
                        break;
                    default:
                        if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                        {
                            activeSubPanel = "Actions";
                        }
                        break;
                }

                continue;
            }

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedActionIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedActionIndex = actions.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    var action = actions[selectedActionIndex];
                    if (action == "Back")
                    {
                        return;
                    }

                    if (action == "Disk space")
                    {
                        activeSubPanel = "Disk space";
                        diskSpace = await LoadDiskSpaceForCloudPcAsync(cloudPc);
                    }
                    else if (action == "Snapshots")
                    {
                        activeSubPanel = "Snapshots";
                        snapshots = await LoadSnapshotsForCloudPcAsync(cloudPc);
                        selectedSnapshotIndex = 0;
                    }
                    else if (action == "Remote action history")
                    {
                        activeSubPanel = "Remote action history";
                        remoteActions = await LoadRemoteActionsForCloudPcAsync(cloudPc);
                        selectedRemoteActionIndex = 0;
                    }
                    else
                    {
                        await InvokeCloudPcActionAsync(cloudPc, action);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    if (activeSubPanel != "Actions")
                    {
                        activeSubPanel = "Actions";
                        break;
                    }
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (key.KeyChar == 'b' || key.KeyChar == 'B' || key.KeyChar == 'q' || key.KeyChar == 'Q')
                    {
                        if (activeSubPanel != "Actions")
                        {
                            activeSubPanel = "Actions";
                            break;
                        }
                        return;
                    }

                    if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private async Task<CloudPcDiskSpace?> LoadDiskSpaceForCloudPcAsync(CloudPcSummary cloudPc)
    {
        var results = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading disk space...", async _ =>
            {
                IReadOnlyList<CloudPcSummary> targets = new[] { cloudPc };
                return await _session.Graph.GetCloudPcDiskSpacesAsync(targets);
            });

        return results.FirstOrDefault();
    }

    private async Task<GraphTableRow?> LoadSignInStatusForCloudPcAsync(CloudPcSummary cloudPc)
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading sign-in status...", async _ => await _session.Graph.GetSignInStatusRowAsync(cloudPc));
    }

    private async Task<GraphTableRow?> LoadLatestSessionForCloudPcAsync(CloudPcSummary cloudPc)
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading connection status...", async _ =>
            {
                try
                {
                    var rows = await _session.Graph.GetConnectionHistoryReportAsync(cloudPc, 1);
                    return rows.FirstOrDefault();
                }
                catch
                {
                    return null;
                }
            });
    }

    private async Task<IReadOnlyList<CloudPcSnapshot>> LoadSnapshotsForCloudPcAsync(CloudPcSummary cloudPc)
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading snapshots...", async _ => await _session.Graph.GetCloudPcSnapshotsAsync(cloudPc));
    }

    private async Task<IReadOnlyList<CloudPcRemoteActionResult>> LoadRemoteActionsForCloudPcAsync(CloudPcSummary cloudPc)
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading remote action history...", async _ => await _session.Graph.GetCloudPcRemoteActionResultsAsync(cloudPc));
    }

    private static void RenderCloudPcDetailLayout(CloudPcSummary cloudPc, string[] actions, int selectedActionIndex, string activeSubPanel, CloudPcDiskSpace? diskSpace, GraphTableRow? signInStatus, GraphTableRow? latestSession, IReadOnlyList<CloudPcSnapshot>? snapshots, int selectedSnapshotIndex, IReadOnlyList<CloudPcRemoteActionResult>? remoteActions, int selectedRemoteActionIndex)
    {
        RenderTopNav();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape($"W365 CLI > Cloud PCs > {cloudPc.Name}")}[/]");
        AnsiConsole.WriteLine();

        var details = new Panel(
            new Rows(
                new Markup(PropertyInline("Name", cloudPc.Name, "grey")),
                new Markup(PropertyInline("Status", StatusMarkup(cloudPc.Status), valueIsMarkup: true)),
                new Markup(PropertyInline("Power state", cloudPc.PowerState ?? "-", "grey")),
                new Markup(PropertyInline("In use", GetInUseStatusMarkup(latestSession), valueIsMarkup: true)),
                new Markup(PropertyInline("Sign-in status", GetSignInStatusValue(signInStatus, "SignInStatus"), "grey")),
                new Markup(PropertyInline("Last sign-in", GetSignInStatusValue(signInStatus, "LastActiveTime"), "grey")),
                new Markup(PropertyInline("Days since sign-in", GetSignInStatusValue(signInStatus, "DaysSinceLastSignIn"), "grey")),
                new Markup(PropertyInline("Type", cloudPc.ProvisioningType ?? "-", "grey")),
                new Markup(PropertyInline("User", cloudPc.EffectiveUserPrincipalName ?? "-", "grey")),
                new Markup(PropertyInline("Service plan", cloudPc.ServicePlanName ?? "-", "grey")),
                new Markup(PropertyInline("Cloud PC ID", cloudPc.Id, "grey"))))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions
            .Select((action, index) => FormatActionLine(action, index == selectedActionIndex));

        var rightPanel = activeSubPanel switch
        {
            "Disk space" => CreateDiskSpaceSubPanel(diskSpace),
            "Snapshots" => CreateSnapshotsSubPanel(snapshots, selectedSnapshotIndex),
            "Remote action history" => CreateRemoteActionsSubPanel(remoteActions, selectedRemoteActionIndex),
            _ => new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
                .Header("Actions")
                .Border(BoxBorder.Rounded)
        };

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, rightPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(rightPanel);
        }

        var hint = activeSubPanel switch
        {
            "Snapshots" => "Up/Down select snapshot | Enter actions | C/N create | R refresh | Esc/B/Q back to actions",
            "Remote action history" => "Up/Down select action | R refresh | Esc/B/Q back to actions",
            "Disk space" => "Esc/B/Q back to actions",
            _ => "Up/Down choose action | Enter run | Esc/B/Q back | P or Ctrl+K command palette"
        };
        AnsiConsole.MarkupLine($"[grey]{hint}[/]");
        RenderStatusBar();
    }

    private static Panel CreateDiskSpaceSubPanel(CloudPcDiskSpace? disk)
    {
        if (disk is null)
        {
            return new Panel("[yellow]Disk space is unavailable for this Cloud PC.[/]")
                .Header("Disk space")
                .Border(BoxBorder.Rounded);
        }

        var rows = new Rows(
            new Markup($"[bold]Free[/] {Markup.Escape(FormatGb(disk.FreeStorageGb))}"),
            new Markup($"[bold]Used[/] {Markup.Escape(FormatGb(disk.UsedStorageGb))}"),
            new Markup($"[bold]Total[/] {Markup.Escape(FormatGb(disk.TotalStorageGb))}"),
            new Markup($"[bold]Percent free[/] {Markup.Escape(disk.PercentFree is null ? "-" : $"{disk.PercentFree}%")}"),
            new Markup($"[bold]Last sync[/] {Markup.Escape(disk.LastSyncDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
            new Markup($"[bold]Status[/]\n{Markup.Escape(disk.Error ?? "Disk data available")}"),
            new Markup($"[bold]Managed device[/]\n{Markup.Escape(disk.ManagedDeviceName ?? "-")}"));

        return new Panel(rows)
            .Header("Disk space")
            .Border(BoxBorder.Rounded);
    }

    private static string GetSignInStatusValue(GraphTableRow? row, string field)
    {
        if (row is null)
        {
            return "-";
        }

        return GetOptionalField(row, field) ?? "-";
    }

    private static string GetInUseStatusMarkup(GraphTableRow? latestSession)
    {
        if (latestSession is null)
        {
            return "[grey]Unknown[/]";
        }

        var sessionEnd = GetField(latestSession, "SessionEndTime");
        return sessionEnd == "-"
            ? "[yellow]In use[/]"
            : "[green]Available[/]";
    }

    private static Panel CreateSnapshotsSubPanel(IReadOnlyList<CloudPcSnapshot>? snapshots, int selectedSnapshotIndex)
    {
        if (snapshots is null)
        {
            return new Panel("[grey]Snapshots have not been loaded yet.[/]")
                .Header("Snapshots")
                .Border(BoxBorder.Rounded);
        }

        if (snapshots.Count == 0)
        {
            return new Panel("[yellow]No snapshots found for this Cloud PC.[/]\n\n[grey]Press C or N to create the first snapshot.[/]")
                .Header("Snapshots")
                .Border(BoxBorder.Rounded);
        }

        var table = new Table()
            .Border(TableBorder.Simple)
            .AddColumn("Status")
            .AddColumn("Type")
            .AddColumn("Created")
            .AddColumn("Expires");

        var visible = snapshots.Take(Math.Max(3, Console.WindowHeight - 18)).ToArray();
        for (var index = 0; index < visible.Length; index++)
        {
            var snapshot = visible[index];
            var selected = index == selectedSnapshotIndex;
            table.AddRow(
                selected ? Selected(Markup.Escape(snapshot.Status ?? "-")) : Markup.Escape(snapshot.Status ?? "-"),
                selected ? Selected(Markup.Escape(snapshot.SnapshotType ?? "-")) : Markup.Escape(snapshot.SnapshotType ?? "-"),
                selected ? Selected(Markup.Escape(snapshot.CreatedDateTime?.ToLocalTime().ToString("g") ?? "-")) : Markup.Escape(snapshot.CreatedDateTime?.ToLocalTime().ToString("g") ?? "-"),
                selected ? Selected(Markup.Escape(snapshot.ExpirationDateTime?.ToLocalTime().ToString("g") ?? "-")) : Markup.Escape(snapshot.ExpirationDateTime?.ToLocalTime().ToString("g") ?? "-"));
        }

        var rows = new Rows(
            new Markup($"[bold]Total[/] {snapshots.Count}"),
            new Markup("[grey]Enter actions | C/N create[/]"),
            table);

        return new Panel(rows)
            .Header("Snapshots")
            .Border(BoxBorder.Rounded);
    }

    private static Panel CreateRemoteActionsSubPanel(IReadOnlyList<CloudPcRemoteActionResult>? remoteActions, int selectedRemoteActionIndex)
    {
        if (remoteActions is null)
        {
            return new Panel("[grey]Remote action history has not been loaded yet.[/]")
                .Header("Remote actions")
                .Border(BoxBorder.Rounded);
        }

        if (remoteActions.Count == 0)
        {
            return new Panel("[yellow]No remote action history was returned for this Cloud PC.[/]")
                .Header("Remote actions")
                .Border(BoxBorder.Rounded);
        }

        var table = new Table()
            .Border(TableBorder.Simple)
            .AddColumn("Action")
            .AddColumn("State")
            .AddColumn("Started")
            .AddColumn("Updated");

        var visible = remoteActions.Take(Math.Max(3, Console.WindowHeight - 18)).ToArray();
        for (var index = 0; index < visible.Length; index++)
        {
            var action = visible[index];
            var selected = index == selectedRemoteActionIndex;
            table.AddRow(
                selected ? Selected(Markup.Escape(action.ActionName ?? "-")) : Markup.Escape(action.ActionName ?? "-"),
                selected ? Selected(Markup.Escape(action.ActionState ?? "-")) : Markup.Escape(action.ActionState ?? "-"),
                selected ? Selected(Markup.Escape(action.StartDateTime?.ToLocalTime().ToString("g") ?? "-")) : Markup.Escape(action.StartDateTime?.ToLocalTime().ToString("g") ?? "-"),
                selected ? Selected(Markup.Escape(action.LastUpdatedDateTime?.ToLocalTime().ToString("g") ?? "-")) : Markup.Escape(action.LastUpdatedDateTime?.ToLocalTime().ToString("g") ?? "-"));
        }

        var selectedAction = remoteActions[Math.Min(selectedRemoteActionIndex, remoteActions.Count - 1)];
        var hasStatusDetail =
            !string.IsNullOrWhiteSpace(selectedAction.StatusCode) ||
            !string.IsNullOrWhiteSpace(selectedAction.StatusMessage);

        var rows = hasStatusDetail
            ? new Rows(
                new Markup($"[bold]Total[/] {remoteActions.Count}"),
                table,
                new Markup($"[bold]Code[/] {Markup.Escape(selectedAction.StatusCode ?? "-")}"),
                new Markup($"[bold]Message[/]\n{Markup.Escape(selectedAction.StatusMessage ?? "-")}"))
            : new Rows(
                new Markup($"[bold]Total[/] {remoteActions.Count}"),
                table);

        return new Panel(rows)
            .Header("Remote actions")
            .Border(BoxBorder.Rounded);
    }

    private async Task<IReadOnlyList<CloudPcSnapshot>> CreateSnapshotAndReloadAsync(CloudPcSummary cloudPc)
    {
        await ConfirmAndRunAsync("Create snapshot", cloudPc.Name, async () => await _session.Graph.CreateSnapshotAsync(cloudPc.Id), "Cloud PC", cloudPc.Name);
        return await LoadSnapshotsForCloudPcAsync(cloudPc);
    }

    private async Task ShowSnapshotActionMenuAsync(CloudPcSummary cloudPc, CloudPcSnapshot snapshot)
    {
        var action = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("[#58a6ff]Snapshot action[/]")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Restore from this snapshot", "Delete this snapshot", "Back"));

        switch (action)
        {
            case "Restore from this snapshot":
                await ConfirmAndRunAsync("Restore", cloudPc.Name, async () => await _session.Graph.RestoreSnapshotAsync(cloudPc.Id, snapshot.SnapshotId), "Cloud PC", cloudPc.Name);
                break;
            case "Delete this snapshot":
                await ConfirmAndRunAsync("Delete snapshot", snapshot.SnapshotId, async () => await _session.Graph.DeleteSnapshotAsync(cloudPc.Id, snapshot.SnapshotId), "Cloud PC", cloudPc.Name);
                break;
        }
    }

    private async Task InvokeCloudPcActionAsync(CloudPcSummary cloudPc, string action)
    {
        switch (action)
        {
            case "Power on":
                if (!IsCloudPcClearlyOff(cloudPc))
                {
                    TimedMessage("[yellow]Power on is only available when the Cloud PC is powered off.[/]");
                    return;
                }
                await ConfirmAndRunAsync("Power on", cloudPc.Name, async () => await _session.Graph.StartCloudPcAsync(cloudPc.Id));
                break;
            case "Resize":
                await ShowResizeAsync(cloudPc);
                break;
            case "Rename":
                await ShowRenameAsync(cloudPc);
                break;
            case "Restart":
                await ConfirmAndRunAsync("Restart", cloudPc.Name, async () => await _session.Graph.RestartCloudPcAsync(cloudPc.Id));
                break;
            case "Sync":
                if (string.IsNullOrWhiteSpace(cloudPc.ManagedDeviceId))
                {
                    TimedMessage("[yellow]This Cloud PC does not include a managed device ID.[/]");
                    return;
                }
                await ConfirmAndRunAsync("Sync", cloudPc.Name, async () => await _session.Graph.SyncManagedDeviceAsync(cloudPc.ManagedDeviceId));
                break;
            case "Reset local admin password":
                if (string.IsNullOrWhiteSpace(cloudPc.ManagedDeviceId))
                {
                    TimedMessage("[yellow]This Cloud PC does not include a managed device ID.[/]");
                    return;
                }
                await ConfirmAndRunAsync("Reset local admin password", cloudPc.Name, async () => await _session.Graph.ResetLocalAdminPasswordAsync(cloudPc.ManagedDeviceId));
                break;
            case "End grace period":
                if (!IsCloudPcInGracePeriod(cloudPc))
                {
                    TimedMessage("[yellow]End grace period is only available while the Cloud PC is in grace period.[/]");
                    return;
                }
                await ConfirmAndRunAsync("End grace period", cloudPc.Name, async () => await _session.Graph.EndCloudPcGracePeriodAsync(cloudPc.Id));
                break;
            case "Reprovision":
                await ConfirmAndRunAsync("Reprovision", cloudPc.Name, async () => await _session.Graph.ReprovisionCloudPcAsync(cloudPc.Id));
                break;
            case "Connection history":
                await ShowConnectionHistoryReportAsync(cloudPc);
                break;
            case "View status details":
                await ShowCloudPcStatusDetailAsync(cloudPc);
                break;
            default:
                TimedMessage($"[yellow]{Markup.Escape(action)} is not implemented in the native CLI yet.[/]");
                break;
        }
    }

    /// <summary>
    /// Shows why a Cloud PC is provisionedWithWarnings/provisionedWithErrors/failed — mirrors the
    /// Intune portal's "View more information" panel, which is built from the Cloud PC's
    /// statusDetail (code, message, retriable, failedAction, rawError, failedPostProvisionSteps).
    /// </summary>
    private async Task ShowCloudPcStatusDetailAsync(CloudPcSummary cloudPc)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Cloud PCs", cloudPc.Name, "Status details");

        CloudPcStatusDetail? detail;
        try
        {
            detail = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading status details...", async _ => await _session.Graph.GetCloudPcStatusDetailAsync(cloudPc.Id));
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "View status details", cloudPc.Name))
            {
                return;
            }

            if (HandleLockedResourceError(ex, "View status details", cloudPc.Name))
            {
                return;
            }

            TimedMessage($"[red]Failed to load status details: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (detail is null)
        {
            TimedMessage("[yellow]No status details are available for this Cloud PC.[/]");
            return;
        }

        AnsiConsole.Clear();
        RenderBreadcrumb("Cloud PCs", cloudPc.Name, "Status details");
        AnsiConsole.MarkupLine($"[#58a6ff]Status details[/] [grey]{Markup.Escape(cloudPc.Name)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[bold]Status[/] {StatusMarkup(cloudPc.Status)}");
        AnsiConsole.MarkupLine($"[bold]Code[/] {Markup.Escape(detail.Code ?? "-")}");
        AnsiConsole.MarkupLine($"[bold]Message[/] {Markup.Escape(detail.Message ?? "-")}");

        if (detail.AdditionalInformation.TryGetValue("failedAction", out var failedAction) && !string.IsNullOrWhiteSpace(failedAction))
        {
            AnsiConsole.MarkupLine($"[bold]Failed action[/] {Markup.Escape(failedAction)}");
        }

        if (detail.AdditionalInformation.TryGetValue("retriable", out var retriable) && !string.IsNullOrWhiteSpace(retriable))
        {
            AnsiConsole.MarkupLine($"[bold]Retriable[/] {Markup.Escape(retriable)}");
        }

        var failedSteps = GetFailedPostProvisionSteps(detail);
        if (failedSteps.Count > 0)
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[bold]The following configurations did not apply successfully:[/]");
            foreach (var step in failedSteps)
            {
                var (title, description) = DescribePostProvisionStep(step);
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(title)}[/]");
                AnsiConsole.MarkupLine($"[grey]{Markup.Escape(description)}[/]");
            }
        }

        if (detail.AdditionalInformation.TryGetValue("rawError", out var rawError) && !string.IsNullOrWhiteSpace(rawError))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[bold]Raw error output[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(rawError)}[/]");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
    }

    private static IReadOnlyList<string> GetFailedPostProvisionSteps(CloudPcStatusDetail detail)
    {
        if (!detail.AdditionalInformation.TryGetValue("failedPostProvisionSteps", out var raw) || string.IsNullOrWhiteSpace(raw))
        {
            return [];
        }

        try
        {
            var steps = JsonSerializer.Deserialize<string[]>(raw);
            return steps ?? [];
        }
        catch (JsonException)
        {
            // Not JSON — treat the whole value as a single step name rather than dropping it.
            return [raw];
        }
    }

    /// <summary>
    /// Maps a Graph "failedPostProvisionSteps" entry to a human-readable title/description,
    /// matching the wording Intune's portal shows in its "View more information" panel. Keys are
    /// normalized (letters/digits only, lowercased) to tolerate the different casing conventions
    /// Graph uses across step names (e.g. "DevicePreparationProfileTimeout" vs the portal's
    /// "devicePreparationProfileProfileTimedout").
    /// </summary>
    private static (string Title, string Description) DescribePostProvisionStep(string step)
    {
        var key = NormalizeStepKey(step);
        return PostProvisionStepInfo.TryGetValue(key, out var info)
            ? info
            : (step, "This configuration step did not apply successfully. The Cloud PC is still accessible, but may be missing this configuration.");
    }

    private static string NormalizeStepKey(string value)
    {
        return new string(value.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
    }

    private static readonly IReadOnlyDictionary<string, (string Title, string Description)> PostProvisionStepInfo =
        new Dictionary<string, (string, string)>
        {
            [NormalizeStepKey("DevicePreparationProfileTimeout")] = (
                "Windows Autopilot device preparation timed out",
                "The Cloud PC was provisioned but timed out before Autopilot Device Preparation completed. Users can access this Cloud PC, but it may be missing important Intune configuration. Check the Windows Autopilot device preparation deployments status report and adjust the Device Preparation Profile and Provisioning Policy timeout values to ensure that provisioning can complete."),
            [NormalizeStepKey("DevicePreparationProfileProfileTimedout")] = (
                "Windows Autopilot device preparation timed out",
                "The Cloud PC was provisioned but timed out before Autopilot Device Preparation completed. Users can access this Cloud PC, but it may be missing important Intune configuration. Check the Windows Autopilot device preparation deployments status report and adjust the Device Preparation Profile and Provisioning Policy timeout values to ensure that provisioning can complete."),
            [NormalizeStepKey("DevicePreparationProfileProfileErrorOccurDuringProvisioning")] = (
                "Windows Autopilot device preparation deployment",
                "The Cloud PC was provisioned but failed to apply Autopilot Device Preparation. Users can access this Cloud PC, but it may be missing important Intune configuration. Check the Windows Autopilot device preparation deployments status report for details."),
            [NormalizeStepKey("DevicePreparationProfileProfileInternalError")] = (
                "Windows Autopilot device preparation deployment",
                "The Cloud PC was provisioned but failed to apply Autopilot Device Preparation. Users can access this Cloud PC, but it may be missing important Intune configuration. We encountered a service error."),
            [NormalizeStepKey("DevicePreparationProfileProfileNotEnabled")] = (
                "Windows Autopilot device preparation deployment",
                "The Cloud PC was provisioned but failed to apply Autopilot Device Preparation. Users can access this Cloud PC, but it may be missing important Intune configuration. Check that the Device Preparation Policy exists, Cloud PCs are running a supported operating system version, and that the device group used in the Device Preparation Profile is configured properly."),
            [NormalizeStepKey("ApplySecurityBaselines")] = (
                "Secure Default Configurations",
                "The secure default configurations for disabling drive, clipboard USB, and printer redirections were not completed for the Cloud PC."),
            [NormalizeStepKey("AutoStartOneDrive")] = (
                "OneDrive configuration",
                "OneDrive couldn't be configured to start automatically when users sign in."),
            [NormalizeStepKey("BlockHighRiskPorts")] = (
                "Blocking High Risk Ports",
                "One or more high risk port(s) couldn't be disabled."),
            [NormalizeStepKey("ChangePageFileLocation")] = (
                "Page file location",
                "The virtual memory page file couldn't be moved to the C drive."),
            [NormalizeStepKey("ConfigOneDrive")] = (
                "OneDrive configuration",
                "The OneDrive client settings and configuration couldn't be applied."),
            [NormalizeStepKey("DepersistRecycleBin")] = (
                "Recycle Bin configuration",
                "Recycle Bin couldn't be disabled."),
            [NormalizeStepKey("DisableBitLocker")] = (
                "BitLocker configuration",
                "BitLocker drive encryption couldn't be disabled."),
            [NormalizeStepKey("DisableRedirections")] = (
                "Local device access",
                "One or more redirection restrictions (clipboard copy/paste, local drives, printers, or USB devices) couldn't be configured."),
            [NormalizeStepKey("DisableReset")] = (
                "Windows reset",
                "The built-in Windows reset option under Settings couldn't be disabled."),
            [NormalizeStepKey("DisableShutdownButton")] = (
                "Start Menu power icons",
                "The shutdown and restart icons on the Start Menu couldn't be hidden."),
            [NormalizeStepKey("DisableTaskManager")] = (
                "Task Manager access",
                "Task Manager access restrictions couldn't be applied."),
            [NormalizeStepKey("DiskExpansion")] = (
                "Disk allocation",
                "The full OS storage, as defined by the assigned license, couldn't be allocated."),
            [NormalizeStepKey("EnableAttackSurfaceReductionRules")] = (
                "Attack surface reduction rules",
                "Microsoft Defender security attack surface reductions rules couldn't be enabled."),
            [NormalizeStepKey("EnableNestedVirtualization")] = (
                "Nested virtualization",
                "Nested virtualization couldn't be enabled on the Cloud PC."),
            [NormalizeStepKey("EnableVirtualizationBasedSecurityFeatures")] = (
                "Advanced security features",
                "Credential Guard and hypervisor-protected code integrity features couldn't be configured."),
            [NormalizeStepKey("EnableWindowsUpdateBeforeInitialLogon")] = (
                "Windows Update pre-logon",
                "Windows Update couldn't be configured to run before the first user logon."),
            [NormalizeStepKey("HideEdgeFirstRunExperience")] = (
                "Microsoft Edge configuration",
                "Microsoft Edge first run experience couldn't be disabled."),
            [NormalizeStepKey("HideTemporaryDrive")] = (
                "Temporary drive",
                "Temporary drive couldn't be hidden from user view in File Explorer."),
            [NormalizeStepKey("InstallCitrixAgent")] = (
                "Citrix HDX Plus installation",
                "Citrix HDX Plus installation or registration was not completed for the Cloud PC. Please inspect the Citrix HDX Plus installation logs for root cause and retry provisioning. The assigned user may connect to the Cloud PC via Remote Desktop."),
            [NormalizeStepKey("LocalAdmin")] = (
                "Local administrator permissions",
                "Administrator permissions on the Cloud PC, as defined by a User Settings policy, couldn't be granted for the user."),
            [NormalizeStepKey("MmdEnrollment")] = (
                "Microsoft Managed Desktop enrollment",
                "Microsoft Managed Desktop enrollment was not completed for the Cloud PC."),
            [NormalizeStepKey("SetCloudPCRegistryKey")] = (
                "Cloud PC registry key",
                "The Cloud PC registry key couldn't be set correctly."),
            [NormalizeStepKey("SetRdpMaxDisconnectionTime")] = (
                "RDP configuration",
                "Maximum disconnection time for Remote Desktop sessions couldn't be configured."),
            [NormalizeStepKey("SetRdpMaxIdleTime")] = (
                "RDP configuration",
                "Maximum idle time for Remote Desktop connections couldn't be configured."),
            [NormalizeStepKey("SetTeamsWebRTCRegistryKey")] = (
                "Teams WebRTC configuration",
                "The WebRTC registry settings for Microsoft Teams couldn't be configured properly."),
            [NormalizeStepKey("SetWindowsUpdatePolicy")] = (
                "Windows Update policy",
                "Windows updates default policies couldn't be configured."),
            [NormalizeStepKey("TimeZoneRedirection")] = (
                "Time zone redirection",
                "The time zone redirection for the device couldn't be configured."),
            [NormalizeStepKey("WindowsAutopatchEnrollment")] = (
                "Windows Autopatch enrollment",
                "Windows Autopatch enrollment was not completed for the Cloud PC."),
            [NormalizeStepKey("WindowsLocalization")] = (
                "Windows language & Region",
                "We're unable to install the selected language for your provisioned Cloud PCs.")
        };

    private async Task ShowConnectionHistoryReportAsync(CloudPcSummary cloudPc)
    {
        const int pageSize = 50;

        await ShowGraphRowsAsync(
            $"Connection history for {cloudPc.Name}",
            async () => await _session.Graph.GetConnectionHistoryReportAsync(cloudPc, pageSize, 0),
            GetConnectionHistoryReportHeader,
            FormatConnectionHistoryReportRow,
            loadMoreAsync: async (skip, top) => await _session.Graph.GetConnectionHistoryReportAsync(cloudPc, top, skip),
            pageBatchSize: pageSize);
    }

    private async Task ShowResizeAsync(CloudPcSummary cloudPc)
    {
        var plans = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading service plans...", async _ => await _session.Graph.GetCloudPcServicePlansAsync());

        if (plans.Count == 0)
        {
            TimedMessage("[yellow]No service plans returned.[/]");
            return;
        }

        var selectedPlanIndex = plans
            .Select((plan, index) => new { plan, index })
            .FirstOrDefault(item => string.Equals(item.plan.Name, cloudPc.ServicePlanName, StringComparison.OrdinalIgnoreCase))
            ?.index ?? 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Cloud PCs", cloudPc.Name, "Resize");
            AnsiConsole.MarkupLine($"[#58a6ff]Resize[/] [grey]{Markup.Escape(cloudPc.Name)}[/]");
            AnsiConsole.MarkupLine($"Current service plan: [grey]{Markup.Escape(cloudPc.ServicePlanName ?? "-")}[/]");
            AnsiConsole.WriteLine();

            RenderServicePlanTable(plans, selectedPlanIndex);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedPlanIndex = Math.Max(0, selectedPlanIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedPlanIndex = Math.Min(plans.Count - 1, selectedPlanIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedPlanIndex = Math.Max(0, selectedPlanIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedPlanIndex = Math.Min(plans.Count - 1, selectedPlanIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedPlanIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedPlanIndex = plans.Count - 1;
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.Enter:
                    var plan = plans[selectedPlanIndex];
                    await ConfirmAndRunAsync(
                        "Resize",
                        $"{cloudPc.Name} to {plan.Name}",
                        async () => await _session.Graph.ResizeCloudPcAsync(cloudPc.Id, plan.Id));

                    return;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private static void RenderServicePlanTable(IReadOnlyList<CloudPcServicePlan> plans, int selectedPlanIndex)
    {
        AnsiConsole.MarkupLine("[#58a6ff]Select target service plan[/]");
        var header = Row("Name", 46, "Type", 12, "vCPU", 6, "RAM", 8, "Storage", 10, "Profile", 10);
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

        var pageSize = Math.Max(8, Math.Min(20, Console.WindowHeight - 14));
        var start = Math.Clamp(selectedPlanIndex - pageSize / 2, 0, Math.Max(0, plans.Count - pageSize));
        var visible = plans.Skip(start).Take(pageSize).ToArray();

        for (var index = 0; index < visible.Length; index++)
        {
            var plan = visible[index];
            var absoluteIndex = start + index;
            var row = Row(
                plan.Name, 46,
                plan.Type ?? "-", 12,
                plan.VCpuCount?.ToString() ?? "-", 6,
                plan.RamGb is null ? "-" : $"{plan.RamGb} GB", 8,
                plan.StorageGb is null ? "-" : $"{plan.StorageGb} GB", 10,
                plan.UserProfileGb is null ? "-" : $"{plan.UserProfileGb} GB", 10);

            var escaped = Markup.Escape(row);
            AnsiConsole.MarkupLine(absoluteIndex == selectedPlanIndex
                ? $"[black on #58a6ff]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter resize | Esc/B/Q back[/]");
    }

    private async Task ShowRenameAsync(CloudPcSummary cloudPc)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Cloud PCs", cloudPc.Name, "Rename");
        AnsiConsole.MarkupLine($"[#58a6ff]Rename[/] [grey]{Markup.Escape(cloudPc.Name)}[/]");
        AnsiConsole.WriteLine();

        var newDisplayName = AnsiConsole.Ask<string>("New Cloud PC display name:");
        if (string.IsNullOrWhiteSpace(newDisplayName))
        {
            TimedMessage("[yellow]Rename cancelled. Display name is required.[/]");
            return;
        }

        await ConfirmAndRunAsync(
            "Rename",
            $"{cloudPc.Name} to {newDisplayName}",
            async () => await _session.Graph.RenameCloudPcAsync(cloudPc.Id, newDisplayName));
    }

    private async Task ConfirmAndRunAsync(string action, string target, Func<Task> operation, string resourceType = "Cloud PC", string? resourceName = null, string requiredPermission = "CloudPC.ReadWrite.All")
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[#58a6ff]{Markup.Escape(action)}[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();

        var confirm = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title($"Submit {Markup.Escape(action)} now?")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Confirm", "Cancel"));

        if (confirm != "Confirm")
        {
            ShowActionResult("Cancelled", action, target, "[yellow]Cancelled.[/]");
            return;
        }

        try
        {
            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Submitting {action}...", async _ => await operation());
            AddActionHistory(action, target, "Submitted", resourceType: resourceType, resourceName: resourceName);
            ShowActionResult("Submitted", action, target, "[green]Submitted.[/]");
        }
        catch (Exception ex)
        {
            AddActionHistory(action, target, "Failed", ex.Message, resourceType, resourceName);
            if (await HandlePermissionErrorAsync(ex, action, target, requiredPermission))
            {
                return;
            }

            if (HandleLockedResourceError(ex, action, target))
            {
                return;
            }

            ShowActionResult("Failed", action, target, "[red]Action failed.[/]", ex.Message);
        }
    }

    private static bool IsPermissionError(Exception ex)
    {
        return ex.Message.Contains("403", StringComparison.Ordinal) ||
            ex.Message.Contains("Forbidden", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("Authorization_RequestDenied", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// When a Graph call fails with a permission-shaped error (403/Forbidden/Authorization_RequestDenied),
    /// shows an actionable recovery screen instead of a bare error, and offers to open the admin
    /// consent page or retry sign-in with a fresh consent prompt. Returns true if it handled the
    /// error (caller should stop further generic error rendering).
    /// </summary>
    private async Task<bool> HandlePermissionErrorAsync(Exception ex, string action, string target, string requiredPermission = "CloudPC.ReadWrite.All")
    {
        if (!IsPermissionError(ex))
        {
            return false;
        }

        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[red]{Markup.Escape(action)} was denied by Microsoft Graph (403 Forbidden).[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(Fit(ex.Message, Math.Max(40, Console.WindowWidth - 4)))}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]This usually means one of two things:[/]");
        AnsiConsole.MarkupLine($"[grey]  1. This app's {Markup.Escape(requiredPermission)} permission hasn't been granted admin consent in your tenant yet.[/]");
        AnsiConsole.MarkupLine("[grey]  2. Your account doesn't have the administrator role required for this specific action.[/]");
        AnsiConsole.WriteLine();

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("What would you like to do?")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Open admin consent page", "Sign in again and re-consent", "Continue"));

        switch (choice)
        {
            case "Open admin consent page":
                OpenUrl(_session.GetAdminConsentUrl());
                TimedMessage("[grey]Opened the admin consent page. Ask a Global or Cloud Application admin to approve it, then retry.[/]");
                break;
            case "Sign in again and re-consent":
                var reconnected = await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Re-authenticating...", async _ => await _session.ReconsentAsync());
                UpdateStatusBarSnapshot();
                TimedMessage(reconnected ? "[green]Re-authenticated. Try the action again.[/]" : "[red]Re-authentication failed.[/]");
                break;
        }

        return true;
    }

    private static bool IsLockedResourceError(Exception ex)
    {
        return ex.Message.Contains("423", StringComparison.Ordinal) ||
            ex.Message.Contains("Locked", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// When a Graph call fails with a 423 Locked error, shows a clear explanation instead of a
    /// bare "Action failed" — 423 almost always means another operation (a pending apply/
    /// reprovision, or a concurrent policy change) is already running against the same resource.
    /// Returns true if it handled the error (caller should stop further generic error rendering).
    /// </summary>
    private static bool HandleLockedResourceError(Exception ex, string action, string target)
    {
        if (!IsLockedResourceError(ex))
        {
            return false;
        }

        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(action)} couldn't run right now — the resource is locked (423 Locked).[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]This usually means another operation is already running against the same resource — most commonly a pending policy apply/reprovision, or a concurrent policy edit.[/]");
        AnsiConsole.MarkupLine("[grey]Wait for that operation to finish, then try again. If this was a policy apply/reprovision, use \"Check reprovision status\" to see when it's done.[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(Fit(ex.Message, Math.Max(40, Console.WindowWidth - 4)))}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
        return true;
    }

    private static void ShowActionResult(string result, string action, string target, string resultMarkup, string? detail = null)
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[#58a6ff]{Markup.Escape(action)}[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(resultMarkup);

        if (string.IsNullOrWhiteSpace(detail))
        {
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Returning...[/]");
            Thread.Sleep(1500);
            return;
        }

        // Long or multi-line detail (e.g. a full Graph error plus the request body we sent) needs
        // to be shown in full and given time to read/copy, rather than Fit-truncated to one line
        // and auto-dismissed after 1.5s.
        var isLong = detail.Contains('\n') || detail.Length > 150;
        if (!isLong)
        {
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(Fit(detail, Math.Max(40, Console.WindowWidth - 4)))}[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Returning...[/]");
            Thread.Sleep(1500);
            return;
        }

        AnsiConsole.WriteLine();
        foreach (var line in detail.Split('\n'))
        {
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(line)}[/]");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
    }

    private async Task ShowCloudAppDetailsAsync(CloudAppSummary app)
    {
        var actions = app.AppStatus?.Equals("published", StringComparison.OrdinalIgnoreCase) == true
            ? new[] { "Unpublish", "Back" }
            : new[] { "Publish", "Back" };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderCloudAppDetailLayout(app, actions, selectedActionIndex);
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedActionIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedActionIndex = actions.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    var action = actions[selectedActionIndex];
                    if (action == "Back")
                    {
                        return;
                    }
                    await InvokeCloudAppActionAsync(app, action);
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                case ConsoleKey.K when key.Modifiers.HasFlag(ConsoleModifiers.Control):
                    await ShowCommandPaletteAsync();
                    break;
                default:
                    if (key.KeyChar == 'b' || key.KeyChar == 'B' || key.KeyChar == 'q' || key.KeyChar == 'Q')
                    {
                        return;
                    }

                    if (key.KeyChar is 'p' or 'P')
                    {
                        await ShowCommandPaletteAsync();
                    }
                    break;
            }
        }
    }

    private async Task InvokeCloudAppActionAsync(CloudAppSummary app, string action)
    {
        switch (action)
        {
            case "Publish":
                await ConfirmAndRunAsync("Publish", app.DisplayName, async () => await _session.Graph.PublishCloudAppAsync(app.Id), "Cloud App", app.DisplayName);
                return;
            case "Unpublish":
                await ConfirmAndRunAsync("Unpublish", app.DisplayName, async () => await _session.Graph.UnpublishCloudAppAsync(app.Id), "Cloud App", app.DisplayName);
                return;
            default:
                return;
        }
    }

    private static void RenderCloudAppDetailLayout(CloudAppSummary app, string[] actions, int selectedActionIndex)
    {
        AnsiConsole.MarkupLine($"[#58a6ff]W365 CLI > Cloud Apps > {Markup.Escape(app.DisplayName)}[/]");
        AnsiConsole.WriteLine();

        var details = new Panel(
            new Rows(
                new Markup($"[bold]Name:[/] {Markup.Escape(app.DisplayName)}"),
                new Markup($"[bold]Status:[/] {AppStatusMarkup(app.AppStatus, 12)}"),
                new Markup($"[bold]Publisher:[/] {Markup.Escape(app.Publisher ?? "-")}"),
                new Markup($"[bold]Discovered app:[/] {Markup.Escape(app.DiscoveredAppName ?? "-")}"),
                new Markup($"[bold]Added:[/] {Markup.Escape(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
                new Markup($"[bold]Published:[/] {Markup.Escape(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
                new Markup($"[bold]Cloud app ID:[/] [grey]{Markup.Escape(app.Id)}[/]")))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));

        var actionsPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
            .Header("Actions")
            .Border(BoxBorder.Rounded);

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, actionsPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(actionsPanel);
        }

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    private sealed record TableChoice<T>(string Label, T? Item, bool IsBack);

    private sealed record SnapshotListItem(CloudPcSummary CloudPc, CloudPcSnapshot Snapshot);

    private sealed record MenuChoice(string Key, string Title, string Description, IReadOnlyList<MenuChoice>? Children = null);

    private sealed record Tip(string Command, string Text);

    private sealed record ActionHistoryItem(string Action, string Target, string ResourceType, string? ResourceName, string Status, DateTimeOffset RequestedAt, string? Detail);

    private sealed record GitHubReleaseInfo(string TagName, string HtmlUrl, DateTimeOffset? PublishedAt, IReadOnlyList<GitHubReleaseAsset> Assets);

    private sealed record GitHubReleaseAsset(string Name, string BrowserDownloadUrl);

    private sealed record LicenseOverviewItem(
        string Family,
        string SkuPartNumbers,
        int Purchased,
        int Assigned,
        int CloudPcCount,
        int DedicatedCloudPcCount,
        int SharedCloudPcCount,
        int ProvisionableCloudPcCount,
        int AvailableCloudPcCount,
        int ActiveSessionLimit,
        int DedicatedUnitsUsed,
        int SharedUnitsUsed,
        int LicenseUnitsUsed,
        int LicenseUnitsLeft,
        IReadOnlyList<CloudPcSummary> CloudPcs,
        IReadOnlyList<ProvisioningPolicySummary> FlexPolicies);

    private sealed record Windows365LicenseInfo(string Family, string PlanKey, string DisplayName);

    private sealed record FlexAccessRow(string GroupName, string UserName, string UserPrincipalName);

    private enum GraphRowSortMode
    {
        None,
        TitleAscending,
        TitleDescending,
        SummaryAscending,
        SummaryDescending
    }

    private enum CloudPcSortMode
    {
        Name,
        Status,
        User,
        ServicePlan
    }

    private enum ProvisioningPolicySortMode
    {
        Name,
        Type,
        Image,
        Join
    }

    private async Task<bool> EnsureConnectedAsync()
    {
        if (_session.IsConnected)
        {
            return true;
        }

        AnsiConsole.MarkupLine("[yellow]Connect to Microsoft Graph first.[/]");
        var connect = AskYesNo("Connect now?");
        if (!connect)
        {
            return false;
        }

        await _session.ConnectAsync();
        await ShowMissingPermissionPromptIfNeededAsync();
        return _session.IsConnected;
    }

    private static void ShowAbout()
    {
        var choices = new[] { "Go to GitHub", "Go to website", "Request feature", "Open issue", "Back" };
        var selectedIndex = 0;
        var topNavIndex = -1;
        while (true)
        {
            AnsiConsole.Clear();
            RenderTopNav("About", topNavIndex);
            var panel = new Panel(new Rows(
                    new Markup($"[bold {AccentColor}]W365 CLI[/] [{MutedColor}]v{GetCurrentVersion()}[/]"),
                    new Markup("[grey]A native .NET keyboard-first experience for Windows 365 Cloud PC workflows.[/]"),
                    new Markup("[grey]This project does not depend on the PowerShell W365CLI module.[/]"),
                    new Markup($"[grey]GitHub:[/] {Markup.Escape(GitHubRepositoryUrl)}"),
                    new Markup($"[grey]Website:[/] {Markup.Escape(ProjectWebsiteUrl)}")))
                .Header("About")
                .Border(BoxBorder.Rounded);
            AnsiConsole.Write(panel);
            AnsiConsole.WriteLine();

            for (var index = 0; index < choices.Length; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on {AccentColor}]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            RenderTopNavAwareHint(topNavIndex, "[grey]Tab top nav | Up/Down move | Enter select | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true, handleTopNavTab: false);

            if (TryHandleTopNavKey(key, ref topNavIndex, currentTabIndex: 1, out var activation))
            {
                switch (activation)
                {
                    case TopNavActivation.Home:
                        throw new NavigateHomeException();
                    case TopNavActivation.Exit:
                        throw new NavigateExitException();
                }

                continue;
            }

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(choices.Length - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = choices.Length - 1;
                    break;
                case ConsoleKey.Enter:
                    switch (choices[selectedIndex])
                    {
                        case "Go to GitHub":
                            OpenUrl(GitHubRepositoryUrl);
                            TimedMessage("[green]Opened GitHub.[/]");
                            break;
                        case "Go to website":
                            OpenUrl(ProjectWebsiteUrl);
                            TimedMessage("[green]Opened website.[/]");
                            break;
                        case "Request feature":
                            OpenUrl(GitHubFeatureUrl);
                            TimedMessage("[green]Opened feature request.[/]");
                            break;
                        case "Open issue":
                            OpenUrl(GitHubIssueUrl);
                            TimedMessage("[green]Opened issue form.[/]");
                            break;
                        case "Back":
                            return;
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return;
                    }
                    break;
            }
        }
    }

    private static void OpenUrl(string url)
    {
        // Process.Start with UseShellExecute=true only reliably opens a browser on Windows.
        // macOS and Linux need their platform-specific launcher commands instead.
        try
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                Process.Start("open", url);
            }
            else
            {
                Process.Start("xdg-open", url);
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[yellow]Couldn't open a browser automatically. Open this URL manually:[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(url)}[/]");
            AnsiConsole.MarkupLine($"[grey]({Markup.Escape(ex.Message)})[/]");
        }
    }

    private static string GetCurrentVersion()
    {
        var version = typeof(W365CliApp).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? typeof(W365CliApp).Assembly.GetName().Version?.ToString(3)
            ?? "0.0.0";
        var metadataIndex = version.IndexOf('+', StringComparison.Ordinal);
        return metadataIndex >= 0 ? version[..metadataIndex] : version;
    }

    private static Version? ParseVersion(string value)
    {
        var normalized = value.Trim().TrimStart('v', 'V');
        var metadataIndex = normalized.IndexOf('+', StringComparison.Ordinal);
        if (metadataIndex >= 0)
        {
            normalized = normalized[..metadataIndex];
        }

        var prereleaseIndex = normalized.IndexOf('-', StringComparison.Ordinal);
        if (prereleaseIndex >= 0)
        {
            normalized = normalized[..prereleaseIndex];
        }

        return Version.TryParse(normalized, out var version) ? version : null;
    }

    private static void Pause()
    {
        TimedMessage("[grey]Returning...[/]");
    }

    private static void TimedMessage(string markup, int milliseconds = 2000)
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(markup);
        Thread.Sleep(milliseconds);
    }

    private static void WaitForBack()
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press Esc, B, or Q to go back...[/]");
        while (true)
        {
            var key = ReadNavigationKey(intercept: true);
            if (key.Key is ConsoleKey.Escape or ConsoleKey.LeftArrow ||
                key.KeyChar is 'b' or 'B' or 'q' or 'Q')
            {
                return;
            }
        }
    }

    /// <summary>
    /// Use this instead of TimedMessage for anything the user needs to actually be able to read
    /// (errors, manual next-steps) right before returning to a screen that will clear/redraw —
    /// a fixed-duration sleep isn't enough time for a multi-line error message.
    /// </summary>
    private static void WaitForAnyKey(string message = "[grey]Press any key to continue...[/]")
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine(message);
        ReadNavigationKey(intercept: true);
    }

    private static ConsoleKeyInfo ReadNavigationKey(bool intercept, bool handleTopNavTab = true)
    {
        var key = Console.ReadKey(intercept);
        if (handleTopNavTab && key.Key == ConsoleKey.Tab)
        {
            throw new NavigateAboutException();
        }

        return key;
    }

    private enum TopNavActivation
    {
        None,
        Home,
        About,
        Exit
    }

    /// <summary>
    /// Shared Home/About/Exit top-nav key handling, used by every screen that owns an interactive
    /// top nav (the main menu and the About screen) so Tab/Left/Right/Down/Escape/Enter behave
    /// identically everywhere. topNavIndex is -1 when the top nav isn't focused; 0/1/2 map to
    /// Home/About/Exit, matching RenderTopNav's layout. currentTabIndex is the tab that matches
    /// the screen currently being shown (0 for the main menu, 1 for the About screen) — the first
    /// Tab press skips straight past it to the next tab, since focusing the tab you're already on
    /// would look like nothing happened. Activation always requires an explicit Enter, so landing
    /// on "Exit" can never be mistaken for having already quit.
    /// </summary>
    private static bool TryHandleTopNavKey(ConsoleKeyInfo key, ref int topNavIndex, int currentTabIndex, out TopNavActivation activation)
    {
        activation = TopNavActivation.None;
        switch (key.Key)
        {
            case ConsoleKey.Tab when topNavIndex < 0:
                topNavIndex = (currentTabIndex + 1) % 3;
                return true;
            case ConsoleKey.Tab when topNavIndex >= 0:
            case ConsoleKey.RightArrow when topNavIndex >= 0:
                topNavIndex = (topNavIndex + 1) % 3;
                return true;
            case ConsoleKey.LeftArrow when topNavIndex >= 0:
                topNavIndex = topNavIndex == 0 ? 2 : topNavIndex - 1;
                return true;
            case ConsoleKey.DownArrow when topNavIndex >= 0:
            case ConsoleKey.Escape when topNavIndex >= 0:
                topNavIndex = -1;
                return true;
            case ConsoleKey.Enter when topNavIndex >= 0:
                activation = topNavIndex switch
                {
                    0 => TopNavActivation.Home,
                    1 => TopNavActivation.About,
                    2 => TopNavActivation.Exit,
                    _ => TopNavActivation.None
                };
                topNavIndex = -1;
                return true;
            default:
                return false;
        }
    }

    private sealed class NavigateHomeException : Exception;

    private sealed class NavigateAboutException : Exception;

    private sealed class NavigateExitException : Exception;
}
