using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
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
        RenderTip();
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
        RenderTransientStatusLineIfAny();
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

    /// <summary>
    /// Connection/tenant/version already appear as dot-status lines in the home header
    /// (<see cref="RenderHeader"/>). Per explicit feedback, no Graph/Tenant status line should
    /// appear on any other screen either — this only prints a transient status message
    /// (e.g. "Disconnected.") when one is active, and nothing otherwise.
    /// </summary>
    private static void RenderTransientStatusLineIfAny()
    {
        if (statusMessage is not null &&
            statusMessageAt is not null &&
            DateTimeOffset.Now - statusMessageAt < TimeSpan.FromSeconds(6))
        {
            AnsiConsole.MarkupLine(statusMessage);
        }
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
    /// The ASCII banner and tip box are never removed — the compact layout only drops blank
    /// spacer lines between them to reclaim a few rows.
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

        const string BannerLine = "██╗    ██╗██████╗  ██████╗ ███████╗     ██████╗██╗     ██╗";
        AnsiConsole.MarkupLine($"[{AccentColor}]{BannerLine}[/]");
        AnsiConsole.MarkupLine($"[{AccentColor}]██║    ██║╚════██╗██╔════╝ ██╔════╝    ██╔════╝██║     ██║[/]");
        AnsiConsole.MarkupLine($"[{AccentColor}]██║ █╗ ██║ █████╔╝███████╗ ███████╗    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine($"[{AccentColor}]██║███╗██║ ╚═══██╗██╔═══██╗╚════██║    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine($"[{AccentColor}]╚███╔███╔╝██████╔╝╚██████╔╝███████║    ╚██████╗███████╗██║[/]");
        AnsiConsole.MarkupLine($"[{AccentColor}] ╚══╝╚══╝ ╚═════╝  ╚═════╝ ╚══════╝     ╚═════╝╚══════╝╚═╝[/]");

        // Right-aligned under the banner's own width (not the full terminal width) — sits at the
        // bottom-right corner of the ASCII art itself rather than stretching across the screen.
        var versionText = $"v{GetCurrentVersion()}";
        var versionPad = new string(' ', Math.Max(0, BannerLine.Length - versionText.Length));
        AnsiConsole.MarkupLine($"{versionPad}[{MutedColor}]{Markup.Escape(versionText)}[/]");
        AnsiConsole.WriteLine();

        var connected = _session.IsConnected;
        var connectionDot = connected ? $"[{GreenColor}]●[/]" : "[red]●[/]";
        var connectionText = connected ? "Connected to Microsoft Graph" : "Not connected to Microsoft Graph";
        AnsiConsole.MarkupLine($"{connectionDot} {connectionText}");

        if (connected)
        {
            var identity = _session.SignedInUserUpn ?? _session.TenantName ?? "signed-in account";
            AnsiConsole.MarkupLine($"[{GreenColor}]●[/] Signed in as: [{TextColor}]{Markup.Escape(identity)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[grey]●[/] Not signed in");
        }

        AnsiConsole.WriteLine();
    }

    private async Task ShowConnectionAsync()
    {
        var choices = _session.IsConnected
            ? new[] { "Disconnect", "Back" }
            : new[] { "Connect", "Back" };

        var choice = PromptChoice(() => RenderHeader(activeNav: null), "[#58a6ff]Connection[/]", choices, "Back");

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

        void RenderContext()
        {
            AnsiConsole.MarkupLine("[yellow]Heads up: this app registration is missing permission(s) in your tenant that some features rely on:[/]");
            foreach (var scope in _session.MissingRequiredScopes)
            {
                AnsiConsole.MarkupLine($"[grey]  - {Markup.Escape(scope)}[/]");
            }
            AnsiConsole.MarkupLine("[grey]Features that use these will fail with 403 Forbidden until a Global or Cloud Application administrator adds and grants consent for them.[/]");
            AnsiConsole.WriteLine();
        }

        var choice = PromptChoice(RenderContext, "How would you like to proceed?", ["Open admin consent page now", "Continue anyway"], "Continue anyway");

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
                case ConsoleKey.LeftArrow:
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

        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            AnsiConsole.MarkupLine($"[#58a6ff]{Markup.Escape(title)}[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");
            for (var index = 0; index < rows.Length; index++)
            {
                var escaped = Markup.Escape(rows[index].Label);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on {AccentColor}]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter select | Esc/B/Q back[/]");

            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = selectedIndex == 0 ? rows.Length - 1 : selectedIndex - 1;
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = (selectedIndex + 1) % rows.Length;
                    break;
                case ConsoleKey.Enter:
                    return rows[selectedIndex].IsBack ? default : rows[selectedIndex].Item;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return default;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return default;
                    }
                    break;
            }
        }
    }

    private static bool Contains(string? value, string filter)
    {
        return value?.Contains(filter, StringComparison.OrdinalIgnoreCase) == true;
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
        return AnsiConsole.Prompt(new TextPrompt<string>("Filter [[Enter blank to clear]]:").AllowEmpty());
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
        RenderTransientStatusLineIfAny();
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

    /// <summary>
    /// App-wide arrow-key selection prompt used instead of Spectre's SelectionPrompt so every
    /// choice screen behaves identically to the rest of the app: Up/Down to move, Enter to
    /// select, and — critically — Esc, Left arrow, B, or Q to cancel back out, matching the
    /// "Esc/B/Q back" convention every other custom-rendered screen already uses. Spectre's own
    /// SelectionPrompt has no way to cancel except navigating to an explicit "Back"/"Cancel"
    /// choice and pressing Enter — that gap was the root cause of "the back button doesn't do
    /// anything" reports on the Connection screen, the reprovision flow, and others.
    /// </summary>
    /// <param name="renderContext">
    /// Redraws whatever context (breadcrumb, explanatory text, etc.) should appear above the
    /// choices on every redraw — called after AnsiConsole.Clear() each time the selection moves.
    /// Pass a no-op if the screen has no context to show above the choices.
    /// </param>
    /// <param name="title">Prompt title, rendered above the choices (markup allowed).</param>
    /// <param name="choices">The selectable choices, in order.</param>
    /// <param name="cancelChoice">
    /// The choice value returned when the user cancels via Esc/Left/B/Q. Must be one of
    /// <paramref name="choices"/> — typically "Back", "Cancel", or "Continue" depending on the
    /// screen's existing convention, so calling code's existing `if (choice == "Back")` checks
    /// keep working unchanged.
    /// </param>
    private static string PromptChoice(Action renderContext, string title, IReadOnlyList<string> choices, string cancelChoice)
    {
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            renderContext();
            AnsiConsole.MarkupLine(title);
            for (var index = 0; index < choices.Count; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on {AccentColor}]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter select | Esc/B/Q back[/]");

            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = selectedIndex == 0 ? choices.Count - 1 : selectedIndex - 1;
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = (selectedIndex + 1) % choices.Count;
                    break;
                case ConsoleKey.Enter:
                    return choices[selectedIndex];
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return cancelChoice;
                default:
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
                    {
                        return cancelChoice;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// Standard yes/no confirmation UI for the whole app — an arrow-key selection prompt rather
    /// than a type-y/n-and-Enter prompt, so every confirmation in the CLI behaves identically.
    /// Esc/Left arrow/B/Q are treated the same as selecting "No" (the safe, non-destructive
    /// default), matching the Esc/B/Q-back convention everywhere else in the app.
    /// </summary>
    private static bool AskYesNo(string question, bool defaultToYes = true)
    {
        var choices = defaultToYes ? new[] { "Yes", "No" } : new[] { "No", "Yes" };
        var choice = PromptChoice(() => { }, question, choices, "No");
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

    private async Task ConfirmAndRunAsync(string action, string target, Func<Task> operation, string resourceType = "Cloud PC", string? resourceName = null, string requiredPermission = "CloudPC.ReadWrite.All")
    {
        var confirm = PromptChoice(
            () =>
            {
                AnsiConsole.MarkupLine($"[#58a6ff]{Markup.Escape(action)}[/]");
                AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
                AnsiConsole.WriteLine();
            },
            $"Submit {Markup.Escape(action)} now?",
            ["Confirm", "Cancel"],
            "Cancel");

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

            if (HandleConflictError(ex, action, target))
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

        void RenderContext()
        {
            AnsiConsole.MarkupLine($"[red]{Markup.Escape(action)} was denied by Microsoft Graph (403 Forbidden).[/]");
            AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(Fit(ex.Message, Math.Max(40, Console.WindowWidth - 4)))}[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]This usually means one of two things:[/]");
            AnsiConsole.MarkupLine($"[grey]  1. This app's {Markup.Escape(requiredPermission)} permission hasn't been granted admin consent in your tenant yet.[/]");
            AnsiConsole.MarkupLine("[grey]  2. Your account doesn't have the administrator role required for this specific action.[/]");
            AnsiConsole.WriteLine();
        }

        var choice = PromptChoice(RenderContext, "What would you like to do?", ["Open admin consent page", "Sign in again and re-consent", "Continue"], "Continue");

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

    private static bool IsConflictError(Exception ex)
    {
        return ex.Message.Contains("409", StringComparison.Ordinal) ||
            ex.Message.Contains("Conflict", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// When a Graph call fails with a 409 Conflict, shows a clear explanation instead of a bare
    /// "Action failed" — this almost always means the Cloud PC's actual state has moved on since
    /// this screen last refreshed (e.g. it powered on/off, or finished a transition) between when
    /// its status was fetched and when the action was submitted, so the action no longer applies to
    /// its current state. This is inherently a staleness race, not a bug to retry blindly against —
    /// refreshing first is what actually resolves it.
    /// </summary>
    private static bool HandleConflictError(Exception ex, string action, string target)
    {
        if (!IsConflictError(ex))
        {
            return false;
        }

        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(action)} couldn't run — the Cloud PC's state doesn't match what this screen last showed (409 Conflict).[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]This usually means the Cloud PC's power/provisioning state changed since this list was loaded — for example it already turned on, is mid-transition, or another action is already in flight.[/]");
        AnsiConsole.MarkupLine("[grey]Refresh (R) to see its current state, then try again.[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(Fit(ex.Message, Math.Max(40, Console.WindowWidth - 4)))}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
        return true;
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

    private sealed record TableChoice<T>(string Label, T? Item, bool IsBack);

    private sealed record MenuChoice(string Key, string Title, string Description, IReadOnlyList<MenuChoice>? Children = null);

    private sealed record Tip(string Command, string Text);

    private sealed record ActionHistoryItem(string Action, string Target, string ResourceType, string? ResourceName, string Status, DateTimeOffset RequestedAt, string? Detail);

    private enum GraphRowSortMode
    {
        None,
        TitleAscending,
        TitleDescending,
        SummaryAscending,
        SummaryDescending
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
