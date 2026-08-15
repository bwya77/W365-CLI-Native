using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

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
                        filter = PromptFilter(filter);
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

        var items = await TryReloadSnapshotsAsync();
        if (items is null)
        {
            return;
        }

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
                    var reloadedAfterAction = await TryReloadSnapshotsAsync();
                    if (reloadedAfterAction is null)
                    {
                        return;
                    }

                    items = reloadedAfterAction;
                    selectedIndex = Math.Min(selectedIndex, Math.Max(0, items.Count - 1));
                    if (items.Count == 0)
                    {
                        TimedMessage("[yellow]No snapshots were returned.[/]");
                        return;
                    }
                    break;
                case ConsoleKey.R:
                    var reloaded = await TryReloadSnapshotsAsync();
                    if (reloaded is null)
                    {
                        return;
                    }

                    items = reloaded;
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
                        filter = PromptFilter(filter);
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

                // Bounded-concurrency fetch instead of one Graph call at a time — for a fleet
                // with dozens of Cloud PCs, serial awaiting was slow and multiplied throttling
                // risk. Cap of 5 keeps this well clear of Graph's throttling thresholds even for
                // very large fleets.
                var perCloudPcResults = await ConcurrencyHelper.MapWithConcurrencyAsync(cloudPcs, maxConcurrency: 5, async cloudPc =>
                {
                    // Graph can 500 for an individual Cloud PC (e.g. one that doesn't support
                    // snapshots) — skip that one instead of failing the whole list, since the
                    // other Cloud PCs' snapshots are still perfectly loadable.
                    try
                    {
                        var snapshots = await _session.Graph.GetCloudPcSnapshotsAsync(cloudPc);
                        return snapshots.Select(snapshot => new SnapshotListItem(cloudPc, snapshot)).ToArray();
                    }
                    catch (HttpRequestException)
                    {
                        return [];
                    }
                });

                return perCloudPcResults
                    .SelectMany(snapshotsForOneCloudPc => snapshotsForOneCloudPc)
                    .OrderByDescending(item => item.Snapshot.CreatedDateTime)
                    .ThenBy(item => item.CloudPc.Name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            });
    }

    /// <summary>
    /// Loads (or reloads) the all-snapshots list with friendly error handling instead of letting
    /// an HttpRequestException (e.g. a Graph 500) crash the whole app. Returns null if the load
    /// failed and the caller should return to the previous screen; otherwise returns the loaded
    /// (possibly empty) list.
    /// </summary>
    private async Task<IReadOnlyList<SnapshotListItem>?> TryReloadSnapshotsAsync()
    {
        try
        {
            return await LoadAllSnapshotsAsync();
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Load snapshots", "Cloud PC snapshots"))
            {
                return null;
            }

            if (HandleLockedResourceError(ex, "Load snapshots", "Cloud PC snapshots"))
            {
                return null;
            }

            AnsiConsole.Clear();
            AnsiConsole.MarkupLine("[red]Failed to load snapshots.[/]");
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            TimedMessage("[grey]Returning...[/]");
            return null;
        }
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
                        filter = PromptFilter(filter);
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

    private async Task<IReadOnlyList<CloudPcSummary>> LoadCloudPcsAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud PCs...", async _ => await _session.Graph.GetCloudPcsAsync());
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
        // Microsoft Graph documents no delete API at all for snapshots returned by
        // retrieveSnapshots (the only listing this app uses) -- automatic, manual, AND retention
        // snapshotType values are every one of them fully service-managed and expire on their own
        // per expirationDateTime. "Imported" snapshots (cloudPCSnapshot: purgeImportedSnapshot) are
        // an entirely separate, unrelated feature this app doesn't implement (a different import
        // flow via cloudPCSnapshot: importSnapshot) -- "imported" isn't even a valid
        // cloudPcSnapshotType value, so a snapshot from this list can never match it. There's
        // simply no "Delete" action to offer here.
        var action = PromptChoice(
            () =>
            {
                AnsiConsole.MarkupLine($"[grey]Cloud PC:[/] {Markup.Escape(cloudPc.Name)}");
                AnsiConsole.MarkupLine("[grey]Cloud PC snapshots are fully service-managed -- there's no delete action; they expire automatically based on retention policy.[/]");
            },
            "[#58a6ff]Snapshot action[/]",
            ["Restore from this snapshot", "Back"],
            "Back");

        if (action == "Restore from this snapshot")
        {
            await ConfirmAndRunAsync("Restore", cloudPc.Name, async () => await _session.Graph.RestoreSnapshotAsync(cloudPc.Id, snapshot.SnapshotId), "Cloud PC", cloudPc.Name);
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

        var newDisplayName = PromptTextCancelable("New Cloud PC display name [[Esc cancel]]:");
        if (string.IsNullOrWhiteSpace(newDisplayName))
        {
            TimedMessage("[yellow]Rename cancelled.[/]");
            return;
        }

        await ConfirmAndRunAsync(
            "Rename",
            $"{cloudPc.Name} to {newDisplayName}",
            async () => await _session.Graph.RenameCloudPcAsync(cloudPc.Id, newDisplayName));
    }

    private sealed record SnapshotListItem(CloudPcSummary CloudPc, CloudPcSnapshot Snapshot);

    private enum CloudPcSortMode
    {
        Name,
        Status,
        User,
        ServicePlan
    }
}
