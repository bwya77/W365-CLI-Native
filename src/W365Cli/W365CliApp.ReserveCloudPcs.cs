using Spectre.Console;

namespace W365Cli;

internal sealed partial class W365CliApp
{
    private const long ReserveLowDaysThresholdInSeconds = 3 * 24 * 60 * 60; // 259200 -- "3 or less days" bucket.
    private const double ReserveDefaultTotalDays = 10.0; // Falls back to the standard 10-day/864000-second reserve budget if totalAllocatedTimeInSeconds is ever missing.

    /// <summary>
    /// Reserve Cloud PCs are a distinct provisioning type (fixed-duration allocations drawn from a
    /// reserve pool, not a normal dedicated/shared Cloud PC) with their own lifecycle data
    /// (reserveDeviceDetail/userDetail/groupDetail) the rest of the Cloud PCs area doesn't surface,
    /// so this gets its own tenant-wide fleet view instead of being folded into Browse Cloud PCs.
    /// </summary>
    private async Task ShowReserveCloudPcsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var cloudPcs = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading reserve Cloud PCs...", async _ => await _session.Graph.GetReserveCloudPcsAsync());

        if (cloudPcs.Count == 0)
        {
            TimedMessage("[yellow]No reserve Cloud PCs were found.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = ReserveCloudPcSortMode.DaysLeft;

        while (true)
        {
            var visibleCloudPcs = SortReserveCloudPcs(FilterReserveCloudPcs(cloudPcs, filter), sortMode);
            if (visibleCloudPcs.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleCloudPcs.Count)
            {
                selectedIndex = visibleCloudPcs.Count - 1;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Cloud PCs", "Reserve Cloud PCs");
            AnsiConsole.Write(CreateReserveCloudPcSummaryPanel(cloudPcs, visibleCloudPcs, filter));
            AnsiConsole.Write(CreateReserveCloudPcTable(visibleCloudPcs, selectedIndex));
            AnsiConsole.MarkupLine($"[grey]Sort: {FormatReserveCloudPcSortMode(sortMode)} | Up/Down move | PgUp/PgDn page | Enter details | / filter | C clear | S sort | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
            RenderStatusBar();

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
                    cloudPcs = await _session.Graph.GetReserveCloudPcsAsync();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextReserveCloudPcSortMode(sortMode);
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
                        filter = PromptFilter(filter);
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'q' or 'Q' or 'b' or 'B')
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

    /// <summary>
    /// "Days left" buckets mirror the low-allocation Graph filters used to flag reserves that are
    /// about to be reclaimed (remainingAllocatedTimeInSeconds le 259200 / eq 0), and "Users with
    /// Cloud PC" mirrors the status eq 'provisioned' filter -- all three computed client-side here
    /// from the already-loaded fleet rather than issuing the separate $count-only Graph calls those
    /// filters were originally sourced from, since every reserve Cloud PC (and its
    /// reserveDeviceDetail/status) is already in memory.
    /// </summary>
    private static Panel CreateReserveCloudPcSummaryPanel(IReadOnlyList<CloudPcSummary> allCloudPcs, IReadOnlyList<CloudPcSummary> visibleCloudPcs, string filter)
    {
        var lowDaysCount = allCloudPcs.Count(pc =>
            pc.ReserveDeviceDetail?.RemainingAllocatedTimeInSeconds is { } remaining &&
            remaining is >= 0 and <= ReserveLowDaysThresholdInSeconds);
        var zeroDaysCount = allCloudPcs.Count(pc => pc.ReserveDeviceDetail?.RemainingAllocatedTimeInSeconds == 0);
        var provisionedCount = allCloudPcs.Count(pc => string.Equals(pc.Status, "provisioned", StringComparison.OrdinalIgnoreCase));

        var rows = new List<Spectre.Console.Rendering.IRenderable>
        {
            new Markup($"[bold]Total[/] {allCloudPcs.Count}   [bold]Visible[/] {visibleCloudPcs.Count}   [bold]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
            new Markup($"[bold]Days left[/] (out of {ReserveDefaultTotalDays:0.#} days)   [orange1]With 3 or less days: {lowDaysCount}[/]   [indianred1]With 0 days: {zeroDaysCount}[/]"),
            new Markup($"[bold]Users with Cloud PC[/]   [darkolivegreen3_1]With provisioned Cloud PC: {provisionedCount}[/]")
        };

        return new Panel(new Rows(rows)).Border(BoxBorder.Rounded).Header("Reserve Cloud PC overview");
    }

    private static Table CreateReserveCloudPcTable(IReadOnlyList<CloudPcSummary> visibleCloudPcs, int selectedIndex)
    {
        const int statusWidth = 20;
        const int nameWidth = 28;
        const int userWidth = 24;
        const int poolWidth = 20;
        const int daysWidth = 12;
        const int expiresWidth = 12;
        const int actionWidth = 22;

        var table = new Table()
            .Title("Reserve Cloud PCs")
            .Border(TableBorder.Rounded)
            .AddColumn(new TableColumn(" ") { Width = 1, NoWrap = true })
            .AddColumn(new TableColumn("Name") { Width = nameWidth, NoWrap = true })
            .AddColumn(new TableColumn("Status") { Width = statusWidth, NoWrap = true })
            .AddColumn(new TableColumn("Assigned to") { Width = userWidth, NoWrap = true })
            .AddColumn(new TableColumn("Pool") { Width = poolWidth, NoWrap = true })
            .AddColumn(new TableColumn("Days left") { Width = daysWidth, NoWrap = true })
            .AddColumn(new TableColumn("License expires") { Width = expiresWidth, NoWrap = true })
            .AddColumn(new TableColumn("Last action") { Width = actionWidth, NoWrap = true });

        if (visibleCloudPcs.Count == 0)
        {
            table.AddRow("-", "-", "[grey]No reserve Cloud PCs match the current filter.[/]", "-", "-", "-", "-", "-");
            return table;
        }

        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 15));
        var start = Math.Max(0, Math.Min(selectedIndex - pageSize / 2, Math.Max(0, visibleCloudPcs.Count - pageSize)));
        var end = Math.Min(visibleCloudPcs.Count - 1, start + pageSize - 1);

        for (var index = start; index <= end; index++)
        {
            var pc = visibleCloudPcs[index];
            var selected = index == selectedIndex;
            var userText = pc.EffectiveUserDisplayName;
            var poolText = pc.GroupDetail?.GroupDisplayName ?? "-";
            var expiresText = pc.ReserveDeviceDetail?.LicenseExpirationDateTime?.ToLocalTime().ToString("d") ?? "-";
            var lastActionText = pc.LastRemoteActionResult is { } action
                ? $"{action.ActionName ?? "-"} ({action.ActionState ?? "-"})"
                : "-";

            table.AddRow(
                selected ? "[black on #58a6ff]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(pc.Name, nameWidth))) : Markup.Escape(Fit(pc.Name, nameWidth)),
                selected ? Selected(Markup.Escape(Fit(pc.Status ?? "unknown", statusWidth))) : StatusMarkup(pc.Status, statusWidth),
                selected ? Selected(Markup.Escape(Fit(userText, userWidth))) : Markup.Escape(Fit(userText, userWidth)),
                selected ? Selected(Markup.Escape(Fit(poolText, poolWidth))) : Markup.Escape(Fit(poolText, poolWidth)),
                selected ? Selected(Markup.Escape(Fit(FormatReserveDaysLeft(pc), daysWidth))) : FormatReserveDaysLeftMarkup(pc, daysWidth),
                selected ? Selected(Markup.Escape(Fit(expiresText, expiresWidth))) : Markup.Escape(Fit(expiresText, expiresWidth)),
                selected ? Selected(Markup.Escape(Fit(lastActionText, actionWidth))) : Markup.Escape(Fit(lastActionText, actionWidth)));
        }

        return table;
    }

    private static string FormatReserveDaysLeft(CloudPcSummary pc)
    {
        var detail = pc.ReserveDeviceDetail;
        if (detail?.RemainingDays is not { } remainingDays)
        {
            return "-";
        }

        var totalDays = detail.TotalDays ?? ReserveDefaultTotalDays;
        return $"{Math.Floor(remainingDays):0}/{Math.Round(totalDays):0}";
    }

    private static string FormatReserveDaysLeftMarkup(CloudPcSummary pc, int width)
    {
        var text = Markup.Escape(Fit(FormatReserveDaysLeft(pc), width));
        var remainingSeconds = pc.ReserveDeviceDetail?.RemainingAllocatedTimeInSeconds;

        if (remainingSeconds == 0)
        {
            return $"[indianred1]{text}[/]";
        }

        if (remainingSeconds is >= 0 and <= ReserveLowDaysThresholdInSeconds)
        {
            return $"[orange1]{text}[/]";
        }

        return $"[grey]{text}[/]";
    }

    internal static IReadOnlyList<CloudPcSummary> FilterReserveCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return cloudPcs;
        }

        return cloudPcs
            .Where(pc =>
                Contains(pc.Name, filter) ||
                Contains(pc.Status, filter) ||
                Contains(pc.EffectiveUserDisplayName, filter) ||
                Contains(pc.EffectiveUserPrincipalName, filter) ||
                Contains(pc.GroupDetail?.GroupDisplayName, filter))
            .ToArray();
    }

    internal static IReadOnlyList<CloudPcSummary> SortReserveCloudPcs(IReadOnlyList<CloudPcSummary> cloudPcs, ReserveCloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            ReserveCloudPcSortMode.Status => cloudPcs.OrderBy(pc => pc.Status, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            ReserveCloudPcSortMode.User => cloudPcs.OrderBy(pc => pc.EffectiveUserDisplayName, StringComparer.OrdinalIgnoreCase).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            ReserveCloudPcSortMode.DaysLeft => cloudPcs.OrderBy(pc => pc.ReserveDeviceDetail?.RemainingAllocatedTimeInSeconds ?? long.MaxValue).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => cloudPcs.OrderBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase).ToArray()
        };
    }

    internal static ReserveCloudPcSortMode NextReserveCloudPcSortMode(ReserveCloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            ReserveCloudPcSortMode.DaysLeft => ReserveCloudPcSortMode.Status,
            ReserveCloudPcSortMode.Status => ReserveCloudPcSortMode.User,
            ReserveCloudPcSortMode.User => ReserveCloudPcSortMode.Name,
            _ => ReserveCloudPcSortMode.DaysLeft
        };
    }

    internal static string FormatReserveCloudPcSortMode(ReserveCloudPcSortMode sortMode)
    {
        return sortMode switch
        {
            ReserveCloudPcSortMode.Status => "status",
            ReserveCloudPcSortMode.User => "user",
            ReserveCloudPcSortMode.DaysLeft => "days left",
            _ => "name"
        };
    }

    internal enum ReserveCloudPcSortMode
    {
        DaysLeft,
        Name,
        Status,
        User
    }
}
