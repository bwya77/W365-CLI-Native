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

        var choices = new[] { "Browse Cloud PCs", "By shared pool", "Reserve Cloud PCs", "Disk space", "Snapshots", "Back" };
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
                        case "By shared pool":
                            await ShowCloudPcsBySharedPoolAsync();
                            break;
                        case "Reserve Cloud PCs":
                            await ShowReserveCloudPcsAsync();
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

    /// <summary>
    /// Reuses the same fleet view (summary panel + table + full Cloud PC actions) already used by
    /// Provisioning > Policies > [policy] > View Cloud PCs, just reached directly from Cloud PCs
    /// without navigating into a specific policy's action menu first. Scoped to sharedByEntraGroup
    /// policies only -- Flex Dedicated (sharedByUser) is a 1:1 allocation (one Cloud PC per group
    /// member), not a shared pool multiple people draw from, so it doesn't belong in a "pool" picker
    /// even though it's technically a Flex/shared provisioning type. Plain Enterprise "dedicated"
    /// policies aren't pools either and are already fully covered by Browse Cloud PCs.
    /// </summary>
    private async Task ShowCloudPcsBySharedPoolAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var policies = await LoadProvisioningPoliciesAsync();
        var sharedPolicies = policies
            .Where(IsSharedByEntraGroupPolicy)
            .OrderBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (sharedPolicies.Length == 0)
        {
            TimedMessage("[yellow]No Flex shared pool provisioning policies were found.[/]");
            return;
        }

        var memberCountsByPolicyId = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading pool membership...", async _ =>
            {
                var pairs = await ConcurrencyHelper.MapWithConcurrencyAsync(sharedPolicies, maxConcurrency: 5, async policy => (policy.Id, Count: await GetPoolMemberCountAsync(policy)));
                return pairs.ToDictionary(pair => pair.Id, pair => pair.Count, StringComparer.OrdinalIgnoreCase);
            });

        while (true)
        {
            var policy = SelectFromTable(
                "Select a Flex shared pool",
                Row("Pool", 44, "Type", 18, "Users sharing", 14),
                sharedPolicies,
                p => Row(
                    p.DisplayName, 44,
                    p.ProvisioningType ?? "-", 18,
                    memberCountsByPolicyId.TryGetValue(p.Id, out var count) && count.HasValue ? count.Value.ToString() : "-", 14));

            if (policy is null)
            {
                return;
            }

            await ShowCloudPcsForProvisioningPolicyAsync(policy);
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

    internal static IReadOnlyList<CloudPcDiskSpace> FilterDiskSpaces(IReadOnlyList<CloudPcDiskSpace> items, string filter)
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
        var bar = BuildDiskUsageBar(disk.UsedStorageGb, disk.TotalStorageGb, 30);
        var panel = new Panel(
            new Rows(
                new Markup($"[bold]Cloud PC:[/] {Markup.Escape(disk.CloudPcName)}"),
                new Markup($"[bold]Managed device:[/] {Markup.Escape(disk.ManagedDeviceName ?? "-")}"),
                new Markup($"[bold]User:[/] {Markup.Escape(disk.AssignedUserUpn ?? "-")}"),
                new Markup($"[bold]Free:[/] {Markup.Escape(FormatGb(disk.FreeStorageGb))}"),
                new Markup($"[bold]Used:[/] {Markup.Escape(FormatGb(disk.UsedStorageGb))}"),
                new Markup($"[bold]Total:[/] {Markup.Escape(FormatGb(disk.TotalStorageGb))}"),
                new Markup(bar),
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

    /// <summary>
    /// Same green/yellow/red-at-90%/75% threshold styling as the Provisioning "User experience
    /// sync" storage bar, reused here so disk space usage reads the same way everywhere it's shown.
    /// Returns unescaped Spectre markup -- callers using the plain Row()+Markup.Escape() list style
    /// must append this AFTER escaping the rest of the row, never pass it through Escape() itself.
    /// </summary>
    private static string BuildDiskUsageBar(double? usedGb, double? totalGb, int width)
    {
        if (totalGb is not > 0 || usedGb is null)
        {
            return $"[grey]{new string('░', width)}[/]";
        }

        var ratio = Math.Clamp(usedGb.Value / totalGb.Value, 0, 1);
        var usedSegments = (int)Math.Round(width * ratio);
        var freeSegments = Math.Max(0, width - usedSegments);
        var barColor = ratio >= 0.9 ? "red" : ratio >= 0.75 ? "yellow" : "green";
        return $"[{barColor}]{new string('█', usedSegments)}[/][grey]{new string('░', freeSegments)}[/]";
    }

    private async Task ShowCloudPcsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var cloudPcs = await LoadCloudPcsWithSignInStatusAsync();

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
                    cloudPcs = await LoadCloudPcsWithSignInStatusAsync();
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

    /// <summary>
    /// Used only by the Browse Cloud PCs screen, which is the one place "In use" needs to be
    /// accurate for every row (both the table column and the currently selected Cloud PC's side
    /// panel) -- NOT by LoadCloudPcsAsync's other callers (Disk space, Connectivity history, etc.),
    /// which don't show "In use" at all and shouldn't pay for the extra per-Cloud-PC network calls
    /// this requires. connectivityResult -- the field the bulk cloudPCs list would otherwise supply
    /// this from -- is confirmed unreliable (always null on a live tenant test regardless of
    /// $select), so this instead bulk-fetches the same real-time sign-in status endpoint the
    /// "Sign-in status" report and Cloud PC details screen already use successfully for every
    /// provisioning type (Enterprise, Flex Dedicated, Flex Shared), then stamps each Cloud PC's
    /// result onto CloudPcSummary.RealTimeSignInStatus.
    /// </summary>
    private async Task<IReadOnlyList<CloudPcSummary>> LoadCloudPcsWithSignInStatusAsync()
    {
        var cloudPcs = await LoadCloudPcsAsync();
        if (cloudPcs.Count == 0)
        {
            return cloudPcs;
        }

        var signInRows = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading real-time sign-in status...", async _ =>
            {
                try
                {
                    return await _session.Graph.GetSignInStatusRowsAsync(cloudPcs);
                }
                catch
                {
                    return Array.Empty<GraphTableRow>();
                }
            });

        var statusByCloudPcId = signInRows
            .Select(row => (
                Id: GetOptionalField(row, "Cloud PC ID"),
                Status: GetOptionalField(row, "SignInStatus"),
                LastActive: GetOptionalField(row, "LastActiveTime")))
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Id))
            .ToDictionary(
                entry => entry.Id!,
                entry => (entry.Status, LastActive: DateTimeOffset.TryParse(entry.LastActive, out var parsed) ? parsed : (DateTimeOffset?)null),
                StringComparer.OrdinalIgnoreCase);

        return cloudPcs
            .Select(pc => statusByCloudPcId.TryGetValue(pc.Id, out var status)
                ? pc with { RealTimeSignInStatus = status.Status, RealTimeLastActiveTime = status.LastActive }
                : pc)
            .ToArray();
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
        grid.AddRow(CreateCloudPcTable(allCloudPcs, visibleCloudPcs, selectedIndex, filter, sideBySide: true), CreateCloudPcSidePanel(selectedCloudPc));

        RenderBreadcrumb("Cloud PCs", "Browse");
        AnsiConsole.Write(CreateCloudPcSummaryPanel(allCloudPcs, visibleCloudPcs, filter));

        var showInUse = allCloudPcs.Any(pc => GetNormalizedInUseStatus(pc) is not null);
        var showUser = Console.WindowWidth >= 105;
        var showServicePlan = Console.WindowWidth >= 135;
        var showDevice = Console.WindowWidth >= 150;
        if (Console.WindowWidth >= GetMinimumSideBySideCloudPcTableWidth(showInUse, showUser, showServicePlan, showDevice))
        {
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(CreateCloudPcTable(allCloudPcs, visibleCloudPcs, selectedIndex, filter, sideBySide: false));
            AnsiConsole.Write(CreateCloudPcSidePanel(selectedCloudPc));
        }
        AnsiConsole.MarkupLine($"[grey]Sort: {FormatCloudPcSortMode(sortMode)} | Up/Down move | PgUp/PgDn page | Enter actions | D disk | N snapshots | Z resize | Y sync | / filter | C clear | S sort | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    /// <summary>
    /// The narrowest terminal width the side-by-side "table + Selected Cloud PC panel" grid layout
    /// can render into without the table needing more width than every column's documented minimum
    /// (see <see cref="GetCloudPcWidths(bool,bool,bool,bool,int,bool)"/>) adds up to -- i.e. the point
    /// below which switching to the stacked (table, then panel, full width) layout instead is the
    /// only way to avoid Spectre silently shrinking a column below what it can cleanly truncate.
    /// </summary>
    internal static int GetMinimumSideBySideCloudPcTableWidth(bool showInUse, bool showUser, bool showServicePlan, bool showDevice = false)
    {
        const int statusMin = 12;
        const int typeMin = 9;
        const int inUseMin = 14;
        const int nameMin = 16;
        const int userMin = 16;
        const int servicePlanMin = 16;
        const int deviceMin = 14;
        const int sidePanelReserve = 40;

        var columnCount = 4 + (showInUse ? 1 : 0) + (showUser ? 1 : 0) + (showServicePlan ? 1 : 0) + (showDevice ? 1 : 0);
        var overhead = (3 * columnCount) + 1;

        return overhead + sidePanelReserve + 1 /* selector */
            + statusMin + typeMin + (showInUse ? inUseMin : 0)
            + nameMin + (showUser ? userMin : 0) + (showServicePlan ? servicePlanMin : 0) + (showDevice ? deviceMin : 0);
    }

    private static Panel CreateCloudPcSummaryPanel(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, string filter, int? poolMemberCount = null)
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

        if (allCloudPcs.Any(pc => GetNormalizedInUseStatus(pc) is not null))
        {
            var inUseCount = allCloudPcs.Count(pc => string.Equals(GetNormalizedInUseStatus(pc), "inUse", StringComparison.OrdinalIgnoreCase));
            var availableCount = allCloudPcs.Count(pc => string.Equals(GetNormalizedInUseStatus(pc), "available", StringComparison.OrdinalIgnoreCase));
            rows.Add(new Markup($"[bold]Shared usage[/] [yellow]In use: {inUseCount}[/]  [green]Available: {availableCount}[/]"));
        }

        // Only set when this fleet view is scoped to a single shared-pool policy (Browse Cloud PCs'
        // call site never passes this) -- the number of people sharing access to this pool, sourced
        // from the policy's assigned group(s), not the Cloud PC count above (a pool with 2 Cloud PCs
        // could still be shared by many more users than that, taking turns).
        if (poolMemberCount is not null)
        {
            rows.Add(new Markup($"[bold]Pool members[/] {poolMemberCount} user(s) sharing this pool"));
        }

        return new Panel(new Rows(rows)).Border(BoxBorder.Rounded).Header("Cloud PC fleet");
    }

    private static Table CreateCloudPcTable(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, int selectedIndex, string filter, bool sideBySide = true)
    {
        var showInUse = allCloudPcs.Any(pc => GetNormalizedInUseStatus(pc) is not null);
        var showUser = Console.WindowWidth >= 105;
        var showServicePlan = Console.WindowWidth >= 135;
        var showDevice = Console.WindowWidth >= 150;
        var widths = GetCloudPcWidths(showInUse, showUser, showServicePlan, sideBySide, showDevice);

        // NoWrap on every sized column keeps Spectre's render width matched exactly to what Fit()
        // already truncated/padded cell text to -- without it, any residual mismatch between our
        // computed widths and what Spectre decides to render wraps overflow onto a second, mostly-
        // blank line instead of truncating with "...", which read as inconsistent gaps between rows.
        var table = new Table()
            .Title("Cloud PCs")
            .Border(TableBorder.Rounded)
            .AddColumn(new TableColumn(" ") { Width = 1, NoWrap = true })
            .AddColumn(new TableColumn("Name") { Width = widths.Name, NoWrap = true })
            .AddColumn(new TableColumn("Status") { Width = widths.Status, NoWrap = true })
            .AddColumn(new TableColumn("Type") { Width = widths.Type, NoWrap = true });

        if (showDevice)
        {
            table.AddColumn(new TableColumn("Device") { Width = widths.Device, NoWrap = true });
        }

        if (showInUse)
        {
            table.AddColumn(new TableColumn("In use") { Width = widths.InUse, NoWrap = true });
        }

        if (showUser)
        {
            table.AddColumn(new TableColumn("User") { Width = widths.User, NoWrap = true });
        }

        if (showServicePlan)
        {
            table.AddColumn(new TableColumn("Service plan") { Width = widths.ServicePlan, NoWrap = true });
        }

        if (visibleCloudPcs.Count == 0)
        {
            var emptyCells = new List<string> { "-", "-", "-", "[grey]No Cloud PCs match the current filter.[/]" };
            if (showDevice)
            {
                emptyCells.Add("-");
            }
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
                selected ? Selected(Markup.Escape(Fit(pc.Name, widths.Name))) : Markup.Escape(Fit(pc.Name, widths.Name)),
                selected ? Selected(Markup.Escape(Fit(pc.Status ?? "unknown", widths.Status))) : StatusMarkup(pc.Status, widths.Status),
                selected ? Selected(Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type))) : Markup.Escape(Fit(pc.ProvisioningType ?? "-", widths.Type))
            };

            if (showDevice)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(pc.ManagedDeviceName ?? "-", widths.Device))) : Markup.Escape(Fit(pc.ManagedDeviceName ?? "-", widths.Device)));
            }

            if (showInUse)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(FormatInUsePlain(pc), widths.InUse))) : FormatInUseMarkup(pc));
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

    /// <summary>
    /// <summary>
    /// Normalizes "in use" into "inUse"/"available"/"unavailable"/null across both data sources this
    /// app has for it. Prefers RealTimeSignInStatus (getRealTimeRemoteConnectionStatus, bulk-fetched
    /// once per Cloud PC when Browse Cloud PCs loads/refreshes -- confirmed accurate for Enterprise,
    /// Flex Dedicated, and Flex Shared alike) over ConnectivityResult (the bulk cloudPCs list's own
    /// field, confirmed unreliable/always-null on a live tenant test). "Unavailable" is Graph's own
    /// real value here -- confirmed directly against a live tenant report dump -- but it means two
    /// different things depending on whether the Cloud PC actually exists: for a genuinely
    /// provisioned Cloud PC, a failed/incomplete real-time check still means it CAN be connected to
    /// (just not signed in right now), i.e. Available; only for notProvisioned Cloud PCs (no VM
    /// exists at all yet) does "Unavailable" mean truly unusable.
    /// </summary>
    internal static string? GetNormalizedInUseStatus(CloudPcSummary pc)
    {
        var signIn = pc.RealTimeSignInStatus;
        if (!string.IsNullOrWhiteSpace(signIn))
        {
            if (signIn.Contains("notsignedin", StringComparison.OrdinalIgnoreCase))
            {
                return "available";
            }

            if (signIn.Contains("signedin", StringComparison.OrdinalIgnoreCase))
            {
                return "inUse";
            }

            if (signIn.Contains("unavailable", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(pc.Status, "notProvisioned", StringComparison.OrdinalIgnoreCase)
                    ? "unavailable"
                    : "available";
            }
        }

        return pc.ConnectivityResult?.Status;
    }

    internal static string FormatInUsePlain(CloudPcSummary pc)
    {
        var status = GetNormalizedInUseStatus(pc);
        if (string.IsNullOrWhiteSpace(status))
        {
            return "-";
        }

        if (string.Equals(status, "inUse", StringComparison.OrdinalIgnoreCase))
        {
            // SharedDeviceDetail.SessionStartDateTime (Frontline/shared-only) is a genuine fixed
            // session-start timestamp. RealTimeLastActiveTime was tried here as an Enterprise
            // fallback, but confirmed wrong on a live tenant: it's a continuously-updating "last
            // seen active" heartbeat, not a session start -- it kept advancing on every refresh
            // (13:48 -> 14:03 -> 14:05 for the same still-signed-in session), which is the opposite
            // of what "since" should mean. Enterprise (dedicated) Cloud PCs have no genuine
            // session-start source available, so just show "In use" with no time for them rather
            // than a fabricated, drifting one.
            var sessionStart = pc.SharedDeviceDetail?.SessionStartDateTime;
            return sessionStart is null
                ? "In use"
                : $"In use since {sessionStart.Value.ToLocalTime():t}";
        }

        if (string.Equals(status, "available", StringComparison.OrdinalIgnoreCase))
        {
            return "Available";
        }

        if (string.Equals(status, "unavailable", StringComparison.OrdinalIgnoreCase))
        {
            return "Unavailable";
        }

        return status;
    }

    internal static string FormatInUseMarkup(CloudPcSummary pc)
    {
        var status = GetNormalizedInUseStatus(pc);
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

        // Browse Cloud PCs now bulk-fetches real-time sign-in status for every Cloud PC up front
        // (LoadCloudPcsWithSignInStatusAsync), so this is normally populated for every row. The
        // fallback text only shows if that bulk fetch itself failed for this specific Cloud PC
        // (e.g. a transient per-item error) and connectivityResult (confirmed unreliable on this
        // list endpoint) has nothing either.
        var inUseLine = GetNormalizedInUseStatus(cloudPc) is not null
            ? new Markup(PropertyInline("In use", FormatInUseMarkup(cloudPc), valueIsMarkup: true))
            : new Markup(PropertyBlock("In use", "See Enter details for live status", "grey"));

        var content = new Rows(
            new Markup(PropertyBlock("Name", cloudPc.Name, "grey")),
            new Markup(PropertyBlock("Device", cloudPc.ManagedDeviceName ?? "-", "grey")),
            new Markup(PropertyInline("Status", StatusMarkup(cloudPc.Status), valueIsMarkup: true)),
            new Markup(PropertyInline("Type", cloudPc.ProvisioningType ?? "-", "grey")),
            new Markup(PropertyBlock("User", cloudPc.EffectiveUserPrincipalName ?? "-", "grey")),
            new Markup(PropertyBlock("Service plan", cloudPc.ServicePlanName ?? "-", "grey")),
            inUseLine,
            new Markup(PropertyBlock("Cloud PC ID", cloudPc.Id, "grey")),
            new Markup(PropertyBlock("Actions", "Enter details, A actions", "grey")));

        return new Panel(content)
            .Header("Selected Cloud PC")
            .Border(BoxBorder.Rounded);
    }

    internal static IReadOnlyList<CloudPcSummary> FilterCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return cloudPcs;
        }

        return cloudPcs
            .Where(pc =>
                Contains(pc.Name, filter) ||
                Contains(pc.ManagedDeviceName, filter) ||
                Contains(pc.Status, filter) ||
                Contains(pc.ProvisioningType, filter) ||
                Contains(pc.EffectiveUserPrincipalName, filter) ||
                Contains(pc.ServicePlanName, filter))
            .ToArray();
    }

    internal static IReadOnlyList<CloudPcSummary> SortCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Status => cloudPcs.OrderBy(pc => pc.Status, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            CloudPcSortMode.User => cloudPcs.OrderBy(pc => pc.EffectiveUserPrincipalName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            CloudPcSortMode.ServicePlan => cloudPcs.OrderBy(pc => pc.ServicePlanName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => cloudPcs.OrderBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray()
        };
    }

    internal static CloudPcSortMode NextCloudPcSortMode(CloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            CloudPcSortMode.Name => CloudPcSortMode.Status,
            CloudPcSortMode.Status => CloudPcSortMode.User,
            CloudPcSortMode.User => CloudPcSortMode.ServicePlan,
            _ => CloudPcSortMode.Name
        };
    }

    internal static string FormatCloudPcSortMode(CloudPcSortMode sortMode)
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

    /// <summary>
    /// Computes exact column widths so the rendered table (borders + padding included) fits within
    /// Console.WindowWidth, and reserves room for the "Selected Cloud PC" side panel that sits next
    /// to this table in a two-column Grid (RenderCloudPcBrowser) once the terminal is wide enough.
    /// This mirrors the same fix applied to the Cloud Apps table: the previous version budgeted
    /// widths independently of Spectre's actual per-column render overhead (n+1 border chars + 2
    /// padding chars per column, Rounded border) and didn't reserve anything for the side panel, so
    /// on some terminal widths a cell (most often "In use" or "Name", whichever happened to end up
    /// wider than Spectre could actually render) would wrap onto a second line while others didn't --
    /// exactly the "inconsistent spacing between rows" symptom, since a wrapped row's second line is
    /// nearly all trailing padding and reads as a gap. CreateCloudPcTable also now sets explicit
    /// Width+NoWrap on every column so any residual mismatch truncates with "..." (already handled
    /// by Fit()) instead of wrapping.
    /// </summary>
    internal static (int Name, int Status, int Type, int User, int ServicePlan, int InUse, int Device) GetCloudPcWidths(bool showInUse, bool showUser, bool showServicePlan, bool sideBySide = true, bool showDevice = false) =>
        GetCloudPcWidths(showInUse, showUser, showServicePlan, sideBySide, Console.WindowWidth, showDevice);

    internal static (int Name, int Status, int Type, int User, int ServicePlan, int InUse, int Device) GetCloudPcWidths(bool showInUse, bool showUser, bool showServicePlan, bool sideBySide, int windowWidth, bool showDevice = false)
    {
        const int statusMax = 24; // Fits the longest real cloudPcStatus value ("provisionedWithWarnings").
        const int statusMin = 12;
        const int typeMax = 18; // Fits the longest real provisioningType value ("sharedByEntraGroup").
        const int typeMin = 9; // Fits "dedicated"/"reserve" in full; only "sharedBy*" truncates below this.
        const int inUseMax = 26;
        const int inUseMin = 14;
        const int nameMin = 16;
        const int userMin = 16;
        const int servicePlanMin = 16;
        const int deviceMin = 14; // Fits "DESKTOP-XXXXXXX" (the default Windows-generated device name length) without truncation.
        const int sidePanelReserve = 40; // "Selected Cloud PC" panel width + Grid column gap -- only spent when rendered side-by-side with it.

        var columnCount = 4 + (showInUse ? 1 : 0) + (showUser ? 1 : 0) + (showServicePlan ? 1 : 0) + (showDevice ? 1 : 0); // selector, status, type, name [+ in use] [+ user] [+ service plan] [+ device]
        var overhead = (3 * columnCount) + 1; // Rounded border: (n+1) border chars + 2 padding chars per column
        var sidePanelCost = sideBySide ? sidePanelReserve : 0;

        // Start every sized column at its preferred (max) width, then shrink Type, In use, and
        // Status -- in that priority order -- before ever shrinking Name/User/Service plan/Device
        // below their minimums. Without this, a narrow terminal (or the side panel eating width)
        // left Status/Type fixed at their preferred size while Name/User were floored at a minimum
        // that still didn't fit -- Spectre then silently rendered every column narrower than
        // requested to make the table fit the real terminal width, which broke NoWrap's
        // truncate-with-"..." behavior and wrapped cell text character-by-character instead (e.g.
        // "sharedByUser" rendering as "sh/ar/ed/B.." down a squashed Type column).
        var status = statusMax;
        var type = typeMax;
        var inUse = showInUse ? inUseMax : 0;
        var minNameBudget = nameMin + (showUser ? userMin : 0) + (showServicePlan ? servicePlanMin : 0) + (showDevice ? deviceMin : 0);

        int RemainingBudget() => windowWidth - overhead - sidePanelCost - 1 /* selector */ - status - type - inUse;

        while (RemainingBudget() < minNameBudget && (type > typeMin || (showInUse && inUse > inUseMin) || status > statusMin))
        {
            if (type > typeMin)
            {
                type--;
            }
            else if (showInUse && inUse > inUseMin)
            {
                inUse--;
            }
            else if (status > statusMin)
            {
                status--;
            }
        }

        // Name gets a straight percentage of `available`, but is capped so it can never eat into
        // the guaranteed minimum reserved for User/Device/Service plan -- without this cap, Name's
        // 32% (rounded up past its own minimum whenever `available` is only just at the combined
        // floor) silently stole width the other columns' minimums needed, so the four dynamic
        // columns' requested widths together exceeded the real terminal width and Spectre
        // force-shrank columns below NoWrap's floor, wrapping cell text character-by-character
        // instead of truncating with "...".
        var available = Math.Max(minNameBudget, RemainingBudget());
        var extraMinSum = (showUser ? userMin : 0) + (showDevice ? deviceMin : 0) + (showServicePlan ? servicePlanMin : 0);
        var noExtras = !showUser && !showServicePlan && !showDevice;
        var name = noExtras
            ? available
            : Math.Min(Math.Max(nameMin, (int)(available * 0.32)), Math.Max(nameMin, available - extraMinSum));

        // Whatever's left after Name is split among the active extra columns: each gets at least
        // its minimum, plus an even share of any surplus beyond the combined minimum. Service plan
        // (computed last) always absorbs the exact remainder, so the four dynamic columns' widths
        // can never sum to more than `available` regardless of which combination is active.
        var extrasBudget = Math.Max(extraMinSum, available - name);
        var extraCount = (showUser ? 1 : 0) + (showDevice ? 1 : 0) + (showServicePlan ? 1 : 0);
        var surplus = Math.Max(0, extrasBudget - extraMinSum);
        var perColumnSurplus = extraCount > 0 ? surplus / extraCount : 0;

        var user = showUser ? userMin + perColumnSurplus : 0;
        var device = showDevice ? deviceMin + perColumnSurplus : 0;
        var servicePlan = showServicePlan ? Math.Max(servicePlanMin, extrasBudget - user - device) : 0;
        return (name, status, type, user, servicePlan, inUse, device);
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
                new Markup(PropertyInline("Device", cloudPc.ManagedDeviceName ?? "-", "grey")),
                new Markup(PropertyInline("Status", StatusMarkup(cloudPc.Status), valueIsMarkup: true)),
                new Markup(PropertyInline("Power state", cloudPc.PowerState ?? "-", "grey")),
                new Markup(PropertyInline("In use", GetInUseStatusMarkup(latestSession, signInStatus), valueIsMarkup: true)),
                new Markup(PropertyInline("Sign-in status", GetSignInStatusValue(signInStatus, "SignInStatus"), "grey")),
                new Markup(PropertyInline("Last sign-in", GetSignInStatusValue(signInStatus, "LastActiveTime"), "grey")),
                new Markup(PropertyInline("Days since sign-in", GetSignInStatusValue(signInStatus, "DaysSinceLastSignIn"), "grey")),
                new Markup(PropertyInline("Provisioned", cloudPc.ProvisionedDateTime?.ToLocalTime().ToString("g") ?? "-", "grey")),
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

        var bar = BuildDiskUsageBar(disk.UsedStorageGb, disk.TotalStorageGb, 24);
        var rows = new Rows(
            new Markup($"[bold]Free[/] {Markup.Escape(FormatGb(disk.FreeStorageGb))}"),
            new Markup($"[bold]Used[/] {Markup.Escape(FormatGb(disk.UsedStorageGb))}"),
            new Markup($"[bold]Total[/] {Markup.Escape(FormatGb(disk.TotalStorageGb))}"),
            new Markup(bar),
            new Markup($"[bold]Percent free[/] {Markup.Escape(disk.PercentFree is null ? "-" : $"{disk.PercentFree}%")}"),
            new Markup($"[bold]Last sync[/] {Markup.Escape(disk.LastSyncDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
            new Markup($"[bold]Status[/]\n{Markup.Escape(disk.Error ?? "Disk data available")}"),
            new Markup($"[bold]Managed device[/]\n{Markup.Escape(disk.ManagedDeviceName ?? "-")}"));

        return new Panel(rows)
            .Header("Disk space")
            .Border(BoxBorder.Rounded);
    }

    /// <summary>
    /// Field values here are either plain text (SignInStatus, DaysSinceLastSignIn) or a UTC ISO
    /// timestamp (LastActiveTime, shown as "Last sign-in") straight from Graph -- previously shown
    /// as-is, unlike "In use since" elsewhere on this same screen, which already converts to local
    /// time. Parses and localizes anything that looks like a timestamp; non-date fields simply fail
    /// the parse and pass through unchanged.
    /// </summary>
    private static string GetSignInStatusValue(GraphTableRow? row, string field)
    {
        if (row is null)
        {
            return "-";
        }

        var value = GetOptionalField(row, field);
        if (string.IsNullOrWhiteSpace(value))
        {
            return "-";
        }

        return DateTimeOffset.TryParse(value, out var parsed)
            ? parsed.ToLocalTime().ToString("g")
            : value;
    }

    /// <summary>
    /// "In use" is normally derived from the Connection History report, but that report is
    /// hardcoded to a "Last 7 days" window (troubleshootConnectionConfigurationOfViewDataTableV1Report
    /// filter) -- if the Cloud PC's last logged session started more than 7 days ago, the query
    /// returns zero rows even if the user is actively signed in right now, which showed as a
    /// misleading "Unknown" (confirmed directly: user was signed in per the real-time sign-in
    /// status report at the same moment this showed Unknown). Falls back to the real-time sign-in
    /// status (getRealTimeRemoteConnectionStatus, already loaded alongside this for the "Sign-in
    /// status" field) when the 7-day history has nothing, since a signed-in session IS in use
    /// regardless of when it started.
    /// </summary>
    private static string GetInUseStatusMarkup(GraphTableRow? latestSession, GraphTableRow? signInStatus)
    {
        if (latestSession is not null)
        {
            var sessionEnd = GetField(latestSession, "SessionEndTime");
            return sessionEnd == "-"
                ? "[yellow]In use[/]"
                : "[green]Available[/]";
        }

        var signIn = signInStatus is null ? null : GetOptionalField(signInStatus, "SignInStatus");
        if (!string.IsNullOrWhiteSpace(signIn))
        {
            return signIn.Contains("notsignedin", StringComparison.OrdinalIgnoreCase)
                ? "[green]Available[/]"
                : signIn.Contains("signedin", StringComparison.OrdinalIgnoreCase)
                    ? "[yellow]In use[/]"
                    : "[grey]Unknown[/]";
        }

        return "[grey]Unknown[/]";
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

        var table = NoWrapColumns(new Table()
            .Border(TableBorder.Simple)
            .AddColumn("Status")
            .AddColumn("Type")
            .AddColumn("Created")
            .AddColumn("Expires"));

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

        var table = NoWrapColumns(new Table()
            .Border(TableBorder.Simple)
            .AddColumn("Action")
            .AddColumn("State")
            .AddColumn("Started")
            .AddColumn("Updated"));

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

        var newDisplayName = PromptTextCancelable("New Cloud PC display name:");
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

    internal enum CloudPcSortMode
    {
        Name,
        Status,
        User,
        ServicePlan
    }
}
