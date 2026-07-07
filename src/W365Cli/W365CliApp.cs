using Spectre.Console;

namespace W365Cli;

internal sealed class W365CliApp
{
    private readonly W365Session _session = new();
    private static readonly List<ActionHistoryItem> ActionHistory = [];
    private static string? statusMessage;
    private static DateTimeOffset? statusMessageAt;

    public async Task<int> RunAsync(string[] args)
    {
        Console.Title = "W365 CLI Native";
        await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Checking cached sign-in...", async _ => await _session.TryRestoreAsync());

        var selectedIndex = 0;
        while (true)
        {
            var menuChoices = GetMainMenuChoices();
            if (selectedIndex >= menuChoices.Count)
            {
                selectedIndex = menuChoices.Count - 1;
            }

            RenderMainMenu(menuChoices, selectedIndex);
            var key = Console.ReadKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(menuChoices.Count - 1, selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = menuChoices.Count - 1;
                    break;
                case ConsoleKey.Enter:
                    if (await ExecuteMainMenuChoiceAsync(menuChoices[selectedIndex]))
                    {
                        return 0;
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
                    break;
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
            case "Usage report":
                await ShowGraphRowsAsync(
                    "Windows 365 Cloud PC usage",
                    async () => await _session.Graph.GetUsageRowsAsync(),
                    GetUsageReportHeader,
                    FormatUsageReportRow,
                    OpenCloudPcFromReportRowAsync);
                return false;
            case "Launch details":
                await ShowGraphRowsAsync(
                    "Windows 365 launch details",
                    async () => await _session.Graph.GetLaunchDetailRowsAsync(),
                    GetLaunchDetailsHeader,
                    FormatLaunchDetailsRow);
                return false;
            case "Service plans":
                await ShowGraphRowsAsync("Windows 365 service plans", _session.Graph.GetServicePlanRowsAsync, GetServicePlansHeader, FormatServicePlanRow);
                return false;
            case "Gallery images":
                await ShowGraphRowsAsync("Windows 365 gallery images", _session.Graph.GetGalleryImageRowsAsync, GetGalleryImagesHeader, FormatGalleryImageRow);
                return false;
            case "Supported regions":
                await ShowGraphRowsAsync("Windows 365 supported regions", _session.Graph.GetSupportedRegionRowsAsync, GetSupportedRegionsHeader, FormatSupportedRegionRow);
                return false;
        }

        switch (choice.Key)
        {
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
                ShowPlaceholderArea("Provisioning", "Provisioning policies and maintenance windows will be wired next.");
                break;
            case "Reports":
                await ShowReportsAsync();
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
            AnsiConsole.MarkupLine("[#4091f2]Cloud PCs[/]");
            AnsiConsole.WriteLine();
            for (var index = 0; index < choices.Length; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on #4091f2]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter open | Esc/B/Q back[/]");
            RenderStatusBar();
            var key = Console.ReadKey(intercept: true);
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
            RenderCompactHeader();
            RenderBreadcrumb("Cloud PCs", "Disk space");
            RenderDiskSpaceTable(items, visibleItems, selectedIndex, filter);
            var key = Console.ReadKey(intercept: true);
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
        AnsiConsole.MarkupLine("[#4091f2]Windows 365 Cloud PC disk space[/]");
        AnsiConsole.MarkupLine($"[grey]Rows: {allItems.Count} | Visible: {visibleItems.Count} | Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}[/]");
        var header = Row("Cloud PC", 34, "Free", 10, "Used", 10, "Total", 10, "Free %", 8, "Last sync", 20);
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");
        AnsiConsole.WriteLine();

        var pageSize = Math.Max(8, Console.WindowHeight - 12);
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
                label = Row(
                    disk.CloudPcName, 34,
                    FormatGb(disk.FreeStorageGb), 10,
                    FormatGb(disk.UsedStorageGb), 10,
                    FormatGb(disk.TotalStorageGb), 10,
                    disk.PercentFree is null ? "-" : $"{disk.PercentFree}%", 8,
                    disk.LastSyncDateTime?.ToLocalTime().ToString("g") ?? "-", 20);
            }

            var escaped = Markup.Escape(label);
            AnsiConsole.MarkupLine(index == selectedIndex
                ? $"[black on #4091f2]> {escaped}[/]"
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
                Contains(item.ManagedDeviceId, filter))
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
            var key = Console.ReadKey(intercept: true);
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
                case ConsoleKey.S:
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
        AnsiConsole.MarkupLine("[#4091f2]Windows 365 Cloud PC snapshots[/]");
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

        var pageSize = Math.Max(8, Console.WindowHeight - 10);
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
                ? $"[black on #4091f2]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter Cloud PC actions | S snapshot actions | / or F filter | C clear | R refresh | Esc/B/Q back[/]");
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
            new("CloudPcs", "Cloud PCs", "Browse, inspect, filter, and act on Cloud PCs"),
            new("Provisioning", "Provisioning", "Provisioning policies and maintenance windows"),
            new("Reports", "Reports", "Usage, connectivity, launch details, report streams"),
            new("CloudApps", "Cloud Apps", "Browse, publish, and unpublish Cloud Apps"),
            new("Catalog", "Catalog", "Service plans, images, regions"),
            new("Tenant", "Tenant settings", "Organization settings, profiles, user settings"),
            new("Connection", "Connection", connectionDescription),
            new("About", "About", "Version and project information"),
            new("Exit", "Exit", "Close W365 CLI Native")
        ];
    }

    private void RenderMainMenuDashboard(IReadOnlyList<MenuChoice> choices)
    {
        var connectionText = _session.IsConnected ? "Connected" : "Not connected";
        var connectionColor = _session.IsConnected ? "green" : "yellow";
        var connectionLight = _session.IsConnected ? "[green1]●[/]" : "[yellow]●[/]";
        var tenantName = _session.TenantName ?? "No tenant selected";
        var tenantId = _session.TenantId ?? "-";

        var dashboard = new Grid();
        dashboard.AddColumn();
        dashboard.AddColumn();
        dashboard.AddColumn();
        dashboard.AddRow(
            new Panel(new Markup($"{connectionLight} [bold {connectionColor}]{connectionText}[/]\n[grey]Status[/]")).Border(BoxBorder.Rounded),
            new Panel(new Markup($"[bold white]{Markup.Escape(Fit(tenantId, 36))}[/]\n[grey]Tenant ID[/]")).Border(BoxBorder.Rounded),
            new Panel(new Markup($"[bold white]{Markup.Escape(Fit(tenantName, 30))}[/]\n[grey]Tenant name[/]")).Border(BoxBorder.Rounded));

        AnsiConsole.Write(dashboard);
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Pick an area below. Esc, B, or Q backs out in deeper screens.[/]");
        AnsiConsole.WriteLine();
    }

    private void RenderMainMenu(IReadOnlyList<MenuChoice> choices, int selectedIndex)
    {
        RenderHeader();
        RenderMainMenuDashboard(choices);
        RenderBreadcrumb("Main menu");

        AnsiConsole.MarkupLine("[#4091f2]Select an area[/]");
        for (var index = 0; index < choices.Count; index++)
        {
            var label = FormatMainMenuChoice(choices[index]);
            AnsiConsole.MarkupLine(index == selectedIndex
                ? $"[black on #4091f2]> {label}[/]"
                : $"  {label}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | Enter open | P or Ctrl+K command palette[/]");
        RenderStatusBar();
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
            AnsiConsole.MarkupLine("[#4091f2]Command palette[/]");
            AnsiConsole.MarkupLine($"[grey]Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}[/]");
            AnsiConsole.WriteLine();

            if (visibleCommands.Count == 0)
            {
                AnsiConsole.MarkupLine("[grey]No commands match the current filter.[/]");
            }
            else
            {
                var pageSize = Math.Max(8, Console.WindowHeight - 8);
                var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visibleCommands.Count - pageSize));
                var visiblePage = visibleCommands.Skip(start).Take(pageSize).ToArray();
                for (var index = 0; index < visiblePage.Length; index++)
                {
                    var absoluteIndex = start + index;
                    var label = FormatMainMenuChoice(visiblePage[index]);
                    AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                        ? $"[black on #4091f2]> {label}[/]"
                        : $"  {label}");
                }
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Type to filter | Up/Down move | Enter run | Backspace edit | Esc/B/Q back[/]");
            var key = Console.ReadKey(intercept: true);
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

    private static string FormatMainMenuChoice(MenuChoice choice)
    {
        return $"[white]{Markup.Escape(Fit(choice.Title, 22))}[/] [grey]{Markup.Escape(choice.Description)}[/]";
    }

    private static void RenderBreadcrumb(params string[] parts)
    {
        var allParts = new[] { "W365 CLI Native" }
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
            new("Reports", "Usage report", "Open Cloud PC usage"),
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

    private void RenderHeader()
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine("[#4091f2]██╗    ██╗██████╗  ██████╗ ███████╗     ██████╗██╗     ██╗[/]");
        AnsiConsole.MarkupLine("[#4091f2]██║    ██║╚════██╗██╔════╝ ██╔════╝    ██╔════╝██║     ██║[/]");
        AnsiConsole.MarkupLine("[#4091f2]██║ █╗ ██║ █████╔╝███████╗ ███████╗    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine("[#4091f2]██║███╗██║ ╚═══██╗██╔═══██╗╚════██║    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine("[#4091f2]╚███╔███╔╝██████╔╝╚██████╔╝███████║    ╚██████╗███████╗██║[/]");
        AnsiConsole.MarkupLine("[#4091f2] ╚══╝╚══╝ ╚═════╝  ╚═════╝ ╚══════╝     ╚═════╝╚══════╝╚═╝[/]");
        AnsiConsole.MarkupLine("[grey]W365 CLI Native v0.1.0 | Bradley Wyatt[/]");
        AnsiConsole.WriteLine();
    }

    private async Task ShowConnectionAsync()
    {
        RenderHeader();

        var choices = _session.IsConnected
            ? new[] { "Disconnect", "Back" }
            : new[] { "Connect", "Back" };

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("[#4091f2]Connection[/]")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices(choices));

        switch (choice)
        {
            case "Connect":
                await _session.ConnectAsync();
                TimedMessage("[grey]Returning...[/]");
                break;
            case "Disconnect":
                await _session.DisconnectAsync();
                TimedMessage("[green]Disconnected.[/]");
                break;
        }
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

    private async Task ShowActionHistoryAsync()
    {
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderActionHistory(selectedIndex);
            var key = Console.ReadKey(intercept: true);
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
        AnsiConsole.MarkupLine("[#4091f2]Action history[/]");
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
            .AddColumn("Cloud PC")
            .AddColumn("Target");
        var pageSize = Math.Max(8, Console.WindowHeight - 10);
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, ActionHistory.Count - pageSize));
        var visible = ActionHistory.Skip(start).Take(pageSize).ToArray();
        foreach (var item in visible.Select((value, index) => new { value, index }))
        {
            var absoluteIndex = start + item.index;
            var selectedMarker = absoluteIndex == selectedIndex ? "[black on #4091f2]>[/]" : " ";
            table.AddRow(
                selectedMarker,
                Markup.Escape(item.value.RequestedAt.ToLocalTime().ToString("t")),
                ActionStatusCell(item.value.Status),
                Markup.Escape(item.value.Action),
                Markup.Escape(item.value.CloudPcName ?? "-"),
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
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter remote actions | C clear | Esc/B/Q back[/]");
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
        var cloudPcs = await LoadCloudPcsAsync();
        var cloudPc = cloudPcs.FirstOrDefault(pc =>
            string.Equals(pc.Name, item.CloudPcName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.ManagedDeviceName, item.CloudPcName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(pc.DisplayName, item.CloudPcName, StringComparison.OrdinalIgnoreCase) ||
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
                    .Title("[#4091f2]Reports[/]")
                    .HighlightStyle(SelectionHighlightStyle())
                    .AddChoices(
                        "Usage",
                        "Connectivity history",
                        "Launch details",
                        "Cloud PC reports",
                        "Back"));

            switch (choice)
            {
                case "Usage":
                    await ShowGraphRowsAsync(
                        "Windows 365 Cloud PC usage",
                        async () => await _session.Graph.GetUsageRowsAsync(),
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
                    .Title("[#4091f2]Tenant settings[/]")
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
            AnsiConsole.MarkupLine("[#4091f2]Catalog[/]");
            AnsiConsole.MarkupLine("[grey]Plans, images, and regions used by Windows 365.[/]");
            AnsiConsole.WriteLine();

            for (var index = 0; index < choices.Length; index++)
            {
                var escaped = Markup.Escape(choices[index]);
                AnsiConsole.MarkupLine(index == selectedIndex
                    ? $"[black on #4091f2]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | Enter open | Esc/B/Q back[/]");
            RenderStatusBar();
            var key = Console.ReadKey(intercept: true);
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
                async _ => await ShowCloudPcDetailsAsync(cloudPc));
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
            AnsiConsole.MarkupLine("[#4091f2]Cloud PC report[/]");
            AnsiConsole.WriteLine();

            var pageSize = Math.Max(8, Console.WindowHeight - 7);
            var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, reportNames.Count - pageSize));
            var visible = reportNames.Skip(start).Take(pageSize).ToArray();
            for (var index = 0; index < visible.Length; index++)
            {
                var absoluteIndex = start + index;
                var escaped = Markup.Escape(visible[index]);
                AnsiConsole.MarkupLine(absoluteIndex == selectedIndex
                    ? $"[black on #4091f2]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter select | Esc/B/Q back[/]");
            var key = Console.ReadKey(intercept: true);
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
        var input = string.Empty;
        while (true)
        {
            AnsiConsole.Clear();
            AnsiConsole.MarkupLine("[#4091f2]Top rows[/]");
            AnsiConsole.MarkupLine("[grey]Enter a positive number, press Enter for 50, or Esc/B/Q to go back.[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.Markup($"Top rows [50]: {Markup.Escape(input)}");

            var key = Console.ReadKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.Enter:
                    if (string.IsNullOrWhiteSpace(input))
                    {
                        return 50;
                    }

                    if (int.TryParse(input, out var top) && top > 0)
                    {
                        return top;
                    }

                    TimedMessage("[yellow]Enter a positive row count.[/]");
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

                    if (char.IsDigit(key.KeyChar))
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
        Func<GraphTableRow, Task>? enterAction = null)
    {
        IReadOnlyList<GraphTableRow> rows;
        try
        {
            rows = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Loading {title}...", async _ => await loader());
        }
        catch (Exception ex)
        {
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
            RenderGraphRows(title, rows, visibleRows, selectedIndex, filter, sortMode, headerFactory, rowFactory);
            var key = Console.ReadKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visibleRows.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
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
                case ConsoleKey.Enter:
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
            AnsiConsole.MarkupLine("[#4091f2]Select Cloud PC for connectivity history[/]");
            var widths = GetConnectivityCloudPcWidths();
            var header = widths.ServicePlan > 0
                ? Row("Name", widths.Name, "Status", widths.Status, "User", widths.User, "Service plan", widths.ServicePlan)
                : Row("Name", widths.Name, "Status", widths.Status, "User", widths.User);
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
            AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

            var pageSize = Math.Max(8, Console.WindowHeight - 8);
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
                    ? $"[black on #4091f2]> {escaped}[/]"
                    : $"  {escaped}");
            }

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter open | Esc/B/Q back[/]");
            var key = Console.ReadKey(intercept: true);
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
        Func<GraphTableRow, string> rowFactory)
    {
        var header = headerFactory();
        RenderBreadcrumb(title);
        AnsiConsole.MarkupLine($"[#4091f2]{Markup.Escape(title)}[/]");
        AnsiConsole.MarkupLine($"[grey]Rows: {allRows.Count} | Visible: {visibleRows.Count} | Filter: {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)} | Sort: {FormatSortMode(sortMode)}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

        var pageSize = Math.Max(8, Console.WindowHeight - 10);
        if (visibleRows.Count == 0)
        {
            AnsiConsole.MarkupLine("[grey]No rows match the current filter.[/]");
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]/ or F filter | C clear | S sort | Esc/B/Q back[/]");
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
                ? $"[black on #4091f2]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter details | / or F filter | C clear | S sort | Esc/B/Q back[/]");
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
            ? Row("Cloud PC", widths.CloudPc, "Status", widths.Status, "Power", widths.Power, "User", widths.User, "Service plan", widths.ServicePlan)
            : Row("Cloud PC", widths.CloudPc, "Status", widths.Status, "Power", widths.Power, "User", widths.User);
    }

    private static string FormatUsageReportRow(GraphTableRow row)
    {
        var widths = GetUsageReportWidths();
        return widths.ServicePlan > 0
            ? Row(
                GetField(row, "Cloud PC"), widths.CloudPc,
                GetField(row, "Status"), widths.Status,
                GetField(row, "Power state"), widths.Power,
                GetField(row, "User"), widths.User,
                GetField(row, "Service plan"), widths.ServicePlan)
            : Row(
                GetField(row, "Cloud PC"), widths.CloudPc,
                GetField(row, "Status"), widths.Status,
                GetField(row, "Power state"), widths.Power,
                GetField(row, "User"), widths.User);
    }

    private static (int CloudPc, int Status, int Power, int User, int ServicePlan) GetUsageReportWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 12;
        const int power = 10;
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
            var key = Console.ReadKey(intercept: true);

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
                        await ShowCloudPcDetailsAsync(visibleCloudPcs[selectedIndex]);
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
            var key = Console.ReadKey(intercept: true);

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
                .Title($"[#4091f2]{Markup.Escape(title)}[/]\n[grey]{Markup.Escape(header)}[/]\n[grey]{Markup.Escape(new string('-', header.Length))}[/]")
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

        RenderCompactHeader();
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
        AnsiConsole.MarkupLine($"[grey]Sort: {FormatCloudPcSortMode(sortMode)} | Up/Down move | PgUp/PgDn page | Enter actions | D disk | N snapshots | Z resize | Y sync | / filter | C clear | S sort | R refresh | Esc back[/]");
        RenderStatusBar();
    }

    private void RenderCompactHeader()
    {
        if (_session.IsConnected)
        {
            var tenantText = _session.TenantName is not null
                ? $"{_session.TenantName} ({_session.TenantId})"
                : _session.TenantId ?? "unknown";
            AnsiConsole.MarkupLine($"[#4091f2]W365 CLI Native[/] [grey]v0.1.0 | Bradley Wyatt[/]   [green]Connected[/] [grey]{Markup.Escape(tenantText)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[#4091f2]W365 CLI Native[/] [grey]v0.1.0 | Bradley Wyatt[/]   [yellow]Not connected[/]");
        }
        AnsiConsole.WriteLine();
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

        var content = new Rows(
            new Markup($"[bold]Total[/] {allCloudPcs.Count}   [bold]Visible[/] {visibleCloudPcs.Count}   [bold]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
            new Markup($"[bold]Status[/] {Markup.Escape(statusSummary)}"),
            new Markup($"[bold]Type[/] {Markup.Escape(typeSummary)}"));

        return new Panel(content).Border(BoxBorder.Rounded).Header("Cloud PC fleet");
    }

    private static Table CreateCloudPcTable(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, int selectedIndex, string filter)
    {
        var widths = GetCloudPcWidths();
        var table = new Table()
            .Title("Cloud PCs")
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Status")
            .AddColumn("Type")
            .AddColumn("Name");

        var showUser = Console.WindowWidth >= 105;
        var showServicePlan = Console.WindowWidth >= 135;

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
                selected ? "[black on #4091f2]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(pc.Status ?? "unknown", 12))) : StatusMarkup(pc.Status),
                selected ? Selected(Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type))) : Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type)),
                selected ? Selected(Markup.Escape(Fit(pc.Name, widths.Name))) : Markup.Escape(Fit(pc.Name, widths.Name))
            };

            if (showUser)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(pc.UserPrincipalName ?? "-", widths.User))) : Markup.Escape(Fit(pc.UserPrincipalName ?? "-", widths.User)));
            }

            if (showServicePlan)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(pc.ServicePlanName ?? "-", widths.ServicePlan))) : Markup.Escape(Fit(pc.ServicePlanName ?? "-", widths.ServicePlan)));
            }

            table.AddRow(row.ToArray());
        }

        return table;
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
            new Markup(PropertyBlock("User", cloudPc.UserPrincipalName ?? "-", "grey")),
            new Markup(PropertyBlock("Service plan", cloudPc.ServicePlanName ?? "-", "grey")),
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

        RenderCompactHeader();
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
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter details | A actions | / filter | C clear | R refresh | Esc back[/]");
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
                selected ? "[black on #4091f2]>[/]" : " ",
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
                Contains(pc.UserPrincipalName, filter) ||
                Contains(pc.ServicePlanName, filter))
            .ToArray();
    }

    private static IReadOnlyList<CloudPcSummary> SortCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Status => cloudPcs.OrderBy(pc => pc.Status, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            CloudPcSortMode.User => cloudPcs.OrderBy(pc => pc.UserPrincipalName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
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

    private static bool Contains(string? value, string filter)
    {
        return value?.Contains(filter, StringComparison.OrdinalIgnoreCase) == true;
    }

    private static string[] GetCloudPcActions(CloudPcSummary cloudPc)
    {
        var actions = new List<string>
        {
            "Remote action history",
            "Disk space",
            "Snapshots",
            "Resize"
        };

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

    private static string StatusMarkup(string? status)
    {
        var text = status ?? "unknown";
        var color = text.ToLowerInvariant() switch
        {
            "provisioned" or "available" or "ready" => "darkolivegreen3_1",
            "provisioning" or "pending" or "inprogress" => "khaki1",
            "failed" or "error" => "indianred1",
            "ingraceperiod" => "plum1",
            _ => "grey"
        };

        return $"[{color}]{Markup.Escape(Fit(text, 12))}[/]";
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
        return $"[black on #4091f2]{escapedText}[/]";
    }

    private static void SetStatus(string markup)
    {
        statusMessage = markup;
        statusMessageAt = DateTimeOffset.Now;
    }

    private static void RenderStatusBar()
    {
        return;
    }

    private static bool IsActionHistoryHotkey(ConsoleKeyInfo key)
    {
        return key.KeyChar is 'h' or 'H';
    }

    private static void AddActionHistory(string action, string target, string status, string? detail = null)
    {
        ActionHistory.Insert(0, new ActionHistoryItem(action, target, InferCloudPcName(target), status, DateTimeOffset.Now, detail));
        if (ActionHistory.Count > 100)
        {
            ActionHistory.RemoveRange(100, ActionHistory.Count - 100);
        }
    }

    private static string? InferCloudPcName(string target)
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
        return new Style(Color.Black, Color.FromHex("#4091f2"));
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
        const int status = 12;
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
            RenderCloudPcDetailLayout(cloudPc, actions, selectedActionIndex, activeSubPanel, diskSpace, snapshots, selectedSnapshotIndex, remoteActions, selectedRemoteActionIndex);
            var key = Console.ReadKey(intercept: true);

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

    private static void RenderCloudPcDetailLayout(CloudPcSummary cloudPc, string[] actions, int selectedActionIndex, string activeSubPanel, CloudPcDiskSpace? diskSpace, IReadOnlyList<CloudPcSnapshot>? snapshots, int selectedSnapshotIndex, IReadOnlyList<CloudPcRemoteActionResult>? remoteActions, int selectedRemoteActionIndex)
    {
        AnsiConsole.MarkupLine($"[#4091f2]W365 CLI Native > Cloud PCs > {Markup.Escape(cloudPc.Name)}[/]");
        AnsiConsole.WriteLine();

        var details = new Panel(
            new Rows(
                new Markup(PropertyInline("Name", cloudPc.Name, "grey")),
                new Markup(PropertyInline("Status", StatusMarkup(cloudPc.Status), valueIsMarkup: true)),
                new Markup(PropertyInline("Power state", cloudPc.PowerState ?? "-", "grey")),
                new Markup(PropertyInline("Type", cloudPc.ProvisioningType ?? "-", "grey")),
                new Markup(PropertyInline("User", cloudPc.UserPrincipalName ?? "-", "grey")),
                new Markup(PropertyInline("Service plan", cloudPc.ServicePlanName ?? "-", "grey")),
                new Markup(PropertyInline("Cloud PC ID", cloudPc.Id, "grey"))))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions
            .Select((action, index) => index == selectedActionIndex
                ? $"[black on #4091f2]> {Markup.Escape(action)}[/]"
                : $"  {Markup.Escape(action)}");

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
            _ => "Up/Down choose action | Enter run | Esc/B/Q back"
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
            new Markup($"[bold]Managed device[/]\n{Markup.Escape(disk.ManagedDeviceName ?? "-")}"));

        return new Panel(rows)
            .Header("Disk space")
            .Border(BoxBorder.Rounded);
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
        await ConfirmAndRunAsync("Create snapshot", cloudPc.Name, async () => await _session.Graph.CreateSnapshotAsync(cloudPc.Id));
        return await LoadSnapshotsForCloudPcAsync(cloudPc);
    }

    private async Task ShowSnapshotActionMenuAsync(CloudPcSummary cloudPc, CloudPcSnapshot snapshot)
    {
        var action = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("[#4091f2]Snapshot action[/]")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Restore from this snapshot", "Delete this snapshot", "Back"));

        switch (action)
        {
            case "Restore from this snapshot":
                await ConfirmAndRunAsync("Restore", cloudPc.Name, async () => await _session.Graph.RestoreSnapshotAsync(cloudPc.Id, snapshot.SnapshotId));
                break;
            case "Delete this snapshot":
                await ConfirmAndRunAsync("Delete snapshot", snapshot.SnapshotId, async () => await _session.Graph.DeleteSnapshotAsync(cloudPc.Id, snapshot.SnapshotId));
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
            default:
                TimedMessage($"[yellow]{Markup.Escape(action)} is not implemented in the native CLI yet.[/]");
                break;
        }
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
            AnsiConsole.MarkupLine($"[#4091f2]Resize[/] [grey]{Markup.Escape(cloudPc.Name)}[/]");
            AnsiConsole.MarkupLine($"Current service plan: [grey]{Markup.Escape(cloudPc.ServicePlanName ?? "-")}[/]");
            AnsiConsole.WriteLine();

            RenderServicePlanTable(plans, selectedPlanIndex);
            var key = Console.ReadKey(intercept: true);

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
        AnsiConsole.MarkupLine("[#4091f2]Select target service plan[/]");
        var header = Row("Name", 46, "Type", 12, "vCPU", 6, "RAM", 8, "Storage", 10, "Profile", 10);
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(header)}[/]");
        AnsiConsole.MarkupLine($"[grey]{Markup.Escape(new string('-', header.Length))}[/]");

        var pageSize = Math.Max(8, Console.WindowHeight - 10);
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
                ? $"[black on #4091f2]> {escaped}[/]"
                : $"  {escaped}");
        }

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter resize | Esc/B/Q back[/]");
    }

    private async Task ShowRenameAsync(CloudPcSummary cloudPc)
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[#4091f2]Rename[/] [grey]{Markup.Escape(cloudPc.Name)}[/]");
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

    private async Task ConfirmAndRunAsync(string action, string target, Func<Task> operation)
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine($"[#4091f2]{Markup.Escape(action)}[/]");
        AnsiConsole.MarkupLine($"Target: [grey]{Markup.Escape(target)}[/]");
        AnsiConsole.WriteLine();

        var confirm = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title($"Submit {Markup.Escape(action)} now?")
                .HighlightStyle(SelectionHighlightStyle())
                .AddChoices("Confirm", "Cancel"));

        if (confirm != "Confirm")
        {
            return;
        }

        try
        {
            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync($"Submitting {action}...", async _ => await operation());
            AddActionHistory(action, target, "Submitted");
        }
        catch (Exception ex)
        {
            AddActionHistory(action, target, "Failed", ex.Message);
        }
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
            var key = Console.ReadKey(intercept: true);

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
                default:
                    if (key.KeyChar == 'b' || key.KeyChar == 'B' || key.KeyChar == 'q' || key.KeyChar == 'Q')
                    {
                        return;
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
                await ConfirmAndRunAsync("Publish", app.DisplayName, async () => await _session.Graph.PublishCloudAppAsync(app.Id));
                return;
            case "Unpublish":
                await ConfirmAndRunAsync("Unpublish", app.DisplayName, async () => await _session.Graph.UnpublishCloudAppAsync(app.Id));
                return;
            default:
                return;
        }
    }

    private static void RenderCloudAppDetailLayout(CloudAppSummary app, string[] actions, int selectedActionIndex)
    {
        AnsiConsole.MarkupLine($"[#4091f2]W365 CLI Native > Cloud Apps > {Markup.Escape(app.DisplayName)}[/]");
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

        var actionLines = actions.Select((action, index) => index == selectedActionIndex
            ? $"[black on #4091f2]> {Markup.Escape(action)}[/]"
            : $"  {Markup.Escape(action)}");

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

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back[/]");
        RenderStatusBar();
    }

    private sealed record TableChoice<T>(string Label, T? Item, bool IsBack);

    private sealed record SnapshotListItem(CloudPcSummary CloudPc, CloudPcSnapshot Snapshot);

    private sealed record MenuChoice(string Key, string Title, string Description);

    private sealed record ActionHistoryItem(string Action, string Target, string? CloudPcName, string Status, DateTimeOffset RequestedAt, string? Detail);

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

    private async Task<bool> EnsureConnectedAsync()
    {
        if (_session.IsConnected)
        {
            return true;
        }

        AnsiConsole.MarkupLine("[yellow]Connect to Microsoft Graph first.[/]");
        var connect = AnsiConsole.Confirm("Connect now?");
        if (!connect)
        {
            return false;
        }

        await _session.ConnectAsync();
        return _session.IsConnected;
    }

    private static void ShowAbout()
    {
        AnsiConsole.MarkupLine("[#4091f2]W365 CLI Native[/]");
        AnsiConsole.MarkupLine("A native .NET rewrite experiment for Windows 365 Cloud PC workflows.");
        AnsiConsole.MarkupLine("This project does not depend on the PowerShell W365CLI module.");
        AnsiConsole.MarkupLine("[grey]GitHub:[/] https://github.com/bwya77/W365-CLI-Native");
        Pause();
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
            var key = Console.ReadKey(intercept: true);
            if (key.Key is ConsoleKey.Escape or ConsoleKey.LeftArrow ||
                key.KeyChar is 'b' or 'B' or 'q' or 'Q')
            {
                return;
            }
        }
    }
}
