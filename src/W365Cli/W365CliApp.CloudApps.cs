using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

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
                        filter = PromptFilter(filter);
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

    private async Task<IReadOnlyList<CloudAppSummary>> LoadCloudAppsAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud Apps...", async _ => await _session.Graph.GetCloudAppsAsync());
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
        grid.AddRow(CreateCloudAppTable(allApps, visibleApps, selectedIndex, filter, sideBySide: true), CreateCloudAppSidePanel(selectedApp));

        RenderBreadcrumb("Cloud Apps", "Browse");
        AnsiConsole.Write(CreateCloudAppSummaryPanel(allApps, visibleApps, filter));

        var showPublisher = Console.WindowWidth >= 100;
        var showDates = Console.WindowWidth >= 150;
        if (Console.WindowWidth >= GetMinimumSideBySideCloudAppTableWidth(showPublisher, showDates))
        {
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(CreateCloudAppTable(allApps, visibleApps, selectedIndex, filter, sideBySide: false));
            AnsiConsole.Write(CreateCloudAppSidePanel(selectedApp));
        }
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter details | A actions | / filter | C clear | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    /// <summary>
    /// The narrowest terminal width the side-by-side "table + Selected Cloud App panel" grid
    /// layout can render into without the table needing more width than every column's documented
    /// minimum (see <see cref="GetCloudAppWidths(bool,bool,bool,int)"/>) adds up to -- mirrors
    /// <see cref="GetMinimumSideBySideCloudPcTableWidth"/> for the same reason: below this width,
    /// switching to the stacked (table, then panel, full width) layout instead is the only way to
    /// avoid Spectre silently shrinking a column below what it can cleanly truncate.
    /// </summary>
    internal static int GetMinimumSideBySideCloudAppTableWidth(bool showPublisher, bool showDates)
    {
        const int statusMin = 10;
        const int nameMin = 16;
        const int publisherMin = 14;
        const int dateMin = 12;
        const int sidePanelReserve = 40;

        var columnCount = 3 + (showPublisher ? 1 : 0) + (showDates ? 2 : 0);
        var overhead = (3 * columnCount) + 1;

        return overhead + sidePanelReserve + 1 /* selector */
            + statusMin + nameMin
            + (showPublisher ? publisherMin : 0)
            + (showDates ? dateMin * 2 : 0);
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

    private static Table CreateCloudAppTable(IReadOnlyList<CloudAppSummary> allApps, IReadOnlyList<CloudAppSummary> visibleApps, int selectedIndex, string filter, bool sideBySide = true)
    {
        var widths = GetCloudAppWidths(Console.WindowWidth >= 100, Console.WindowWidth >= 150, sideBySide);

        // NoWrap on every sized column is essential here: Spectre's default column overflow mode
        // wraps text that doesn't fit the column Spectre itself decides to render (which can be
        // narrower than the exact width we PadRight() our cell text to in Fit()). Because Fit()'s
        // padding is nearly all trailing spaces, a wrapped cell's second line is effectively blank
        // -- exactly the empty-looking rows seen when this table's widths were made to scale with
        // a very wide terminal without NoWrap. Setting Width+NoWrap explicitly makes each column's
        // render width match Fit()'s truncation point exactly, so overflow is cropped with "..."
        // (already handled by Fit()) instead of wrapping onto a blank continuation line.
        var table = new Table()
            .Title("Cloud Apps")
            .Border(TableBorder.Rounded)
            .AddColumn(new TableColumn(" ") { Width = 1, NoWrap = true })
            .AddColumn(new TableColumn("Status") { Width = widths.Status, NoWrap = true })
            .AddColumn(new TableColumn("Name") { Width = widths.Name, NoWrap = true });

        var showPublisher = Console.WindowWidth >= 100;
        var showDates = Console.WindowWidth >= 150;
        if (showPublisher)
        {
            table.AddColumn(new TableColumn("Publisher") { Width = widths.Publisher, NoWrap = true });
        }
        if (showDates)
        {
            table.AddColumn(new TableColumn("Published") { Width = widths.Published, NoWrap = true });
            table.AddColumn(new TableColumn("Added") { Width = widths.Added, NoWrap = true });
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

    /// <summary>
    /// Computes exact column widths so the rendered table (borders + padding included) fits
    /// within the given window width -- previously this budgeted only a flat "-4"/"-2" fudge
    /// factor (copied loosely from GetCloudPcWidths' simpler 3-column split), which under-
    /// accounted for a Rounded-border table's real overhead (n+1 border chars + 2 padding chars
    /// per column) once the fixed 104/132 cap was removed. That under-accounting let Name/Publisher
    /// grow wider than what Spectre could actually render, and since Spectre wraps overflow by
    /// default (and Fit() pads with trailing spaces), the wrapped continuation showed up as
    /// blank-looking rows. Now paired with explicit Width+NoWrap on each column in
    /// CreateCloudAppTable as a second layer of protection against any residual rounding.
    ///
    /// Status/Published/Added are shrunk (in that priority order) toward their minimums before
    /// Name/Publisher ever go below theirs, and the whole computed total is kept within
    /// <paramref name="windowWidth"/> -- otherwise, in a side-by-side layout with dates shown
    /// (Console.WindowWidth around 150, right where "show dates" turns on), the same
    /// character-by-character column-squishing bug fixed for the Cloud PCs table could occur here
    /// too, since the previous flat "Math.Max(48, ...)" floor could push the computed total past
    /// the real terminal width once the side panel and every optional column were all showing at
    /// once. <paramref name="sideBySide"/> should be false when the table is rendered full-width
    /// with no adjacent side panel (e.g. a narrow terminal, or a caller with no panel at all), so
    /// it doesn't reserve space for a panel that isn't actually there.
    /// </summary>
    internal static (int Status, int Name, int Publisher, int Published, int Added) GetCloudAppWidths(bool showPublisher, bool showDates, bool sideBySide = true) =>
        GetCloudAppWidths(showPublisher, showDates, sideBySide, Console.WindowWidth);

    internal static (int Status, int Name, int Publisher, int Published, int Added) GetCloudAppWidths(bool showPublisher, bool showDates, bool sideBySide, int windowWidth)
    {
        const int statusMax = 12;
        const int statusMin = 10;
        const int dateMax = 18;
        const int dateMin = 12;
        const int nameMin = 16;
        const int nameMax = 60;
        const int publisherMin = 14;
        const int publisherMax = 40;
        const int sidePanelReserve = 40; // "Selected Cloud App" panel width + Grid column gap -- only spent when rendered side-by-side with it.

        var columnCount = 3 + (showPublisher ? 1 : 0) + (showDates ? 2 : 0); // selector, status, name [+ publisher] [+ published, added]
        var overhead = (3 * columnCount) + 1; // Rounded border: (n+1) border chars + 2 padding chars per column
        var sidePanelCost = sideBySide ? sidePanelReserve : 0;

        var status = statusMax;
        var published = showDates ? dateMax : 0;
        var added = showDates ? dateMax : 0;
        var minNameBudget = nameMin + (showPublisher ? publisherMin : 0);

        int RemainingBudget() => windowWidth - overhead - sidePanelCost - 1 /* selector */ - status - published - added;

        // Shrink Published/Added first (a truncated date is still readable with "..."), then
        // Status, before ever touching Name/Publisher's minimums.
        while (RemainingBudget() < minNameBudget && (
            (showDates && published > dateMin) || (showDates && added > dateMin) || status > statusMin))
        {
            if (showDates && published > dateMin)
            {
                published--;
            }
            else if (showDates && added > dateMin)
            {
                added--;
            }
            else if (status > statusMin)
            {
                status--;
            }
        }

        var available = Math.Max(minNameBudget, RemainingBudget());

        // Split by allocating each column's minimum first, then handing out any slack beyond that
        // proportionally (58/42) and capping at each column's maximum -- capping only ever removes
        // width, so name + publisher can never exceed `available` here, unlike clamping each one
        // independently with Math.Max(...min...), which could round both up past their minimum at
        // once and silently push the total over budget right at the narrowest supported width.
        var slack = Math.Max(0, available - minNameBudget);
        var name = showPublisher
            ? Math.Min(nameMax, nameMin + (int)(slack * 0.58))
            : Math.Clamp(available, nameMin, nameMax + publisherMax);
        var publisher = showPublisher ? Math.Min(publisherMax, publisherMin + (slack - (int)(slack * 0.58))) : 0;
        return (status, name, publisher, published, added);
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
}
