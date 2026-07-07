using Spectre.Console;

namespace W365Cli;

internal sealed class W365CliApp
{
    private readonly W365Session _session = new();

    public async Task<int> RunAsync(string[] args)
    {
        Console.Title = "W365 CLI Native";
        await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Checking cached sign-in...", async _ => await _session.TryRestoreAsync());

        while (true)
        {
            RenderHeader();

            var menuChoices = GetMainMenuChoices();
            RenderMainMenuDashboard(menuChoices);
            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<MenuChoice>()
                    .Title("[cyan]Select an area[/]")
                    .PageSize(12)
                    .UseConverter(FormatMainMenuChoice)
                    .AddChoices(menuChoices));

            switch (choice.Key)
            {
                case "Connection":
                    await ShowConnectionAsync();
                    break;
                case "CloudPcs":
                    await ShowCloudPcsAsync();
                    break;
                case "CloudApps":
                    await ShowCloudAppsAsync();
                    break;
                case "Provisioning":
                    ShowPlaceholderArea("Provisioning", "Provisioning policies and maintenance windows will be wired next.");
                    break;
                case "Reports":
                    ShowPlaceholderArea("Reports", "Usage, connectivity, launch details, and report streams will be wired next.");
                    break;
                case "Catalog":
                    ShowPlaceholderArea("Catalog", "Service plans, images, supported regions, and licensing will be wired next.");
                    break;
                case "Tenant":
                    ShowPlaceholderArea("Tenant settings", "Organization settings, setting profiles, and user settings will be wired next.");
                    break;
                case "About":
                    ShowAbout();
                    break;
                case "Exit":
                    return 0;
            }
        }
    }

    private IReadOnlyList<MenuChoice> GetMainMenuChoices()
    {
        var connectionDescription = _session.IsConnected
            ? "Disconnect Microsoft Graph session"
            : "Connect to Microsoft Graph";

        return
        [
            new("CloudPcs", "Cloud operations", "Cloud PCs", "Browse, inspect, filter, and act on Cloud PCs", "cyan"),
            new("CloudApps", "App delivery", "Cloud Apps", "Browse, publish, and unpublish Cloud Apps", "deepskyblue1"),
            new("Provisioning", "Provisioning", "Policies", "Provisioning policies and maintenance windows", "mediumpurple1"),
            new("Reports", "Insights", "Reports", "Usage, connectivity, launch details, report streams", "steelblue1"),
            new("Catalog", "Catalog", "Plans and images", "Service plans, images, regions, licensing", "darkseagreen2"),
            new("Tenant", "Configuration", "Tenant settings", "Organization settings, profiles, user settings", "khaki1"),
            new("Connection", "Session", "Connection", connectionDescription, _session.IsConnected ? "green" : "yellow"),
            new("About", "Help", "About", "Version and project information", "grey"),
            new("Exit", "System", "Exit", "Close W365 CLI Native", "grey")
        ];
    }

    private void RenderMainMenuDashboard(IReadOnlyList<MenuChoice> choices)
    {
        var connectionText = _session.IsConnected ? "Connected" : "Not connected";
        var connectionColor = _session.IsConnected ? "green" : "yellow";
        var tenantText = _session.TenantName ?? _session.TenantId ?? "No tenant selected";
        var areaCount = choices.Count(choice => choice.Key is not "About" and not "Exit");

        var dashboard = new Grid();
        dashboard.AddColumn();
        dashboard.AddColumn();
        dashboard.AddColumn();
        dashboard.AddRow(
            new Panel(new Markup($"[bold {connectionColor}]{connectionText}[/]\n[grey]Graph session[/]")).Border(BoxBorder.Rounded),
            new Panel(new Markup($"[bold cyan]{areaCount}[/]\n[grey]Navigation areas[/]")).Border(BoxBorder.Rounded),
            new Panel(new Markup($"[bold white]{Markup.Escape(Fit(tenantText, 30))}[/]\n[grey]Tenant context[/]")).Border(BoxBorder.Rounded));

        AnsiConsole.Write(dashboard);
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Pick an area below. Esc, B, or Q backs out in deeper screens.[/]");
        AnsiConsole.WriteLine();
    }

    private static string FormatMainMenuChoice(MenuChoice choice)
    {
        return $"[{choice.Accent}]{Markup.Escape(Fit(choice.Category, 18))}[/] {Markup.Escape(Fit(choice.Title, 20))} [grey]{Markup.Escape(choice.Description)}[/]";
    }

    private void RenderHeader()
    {
        AnsiConsole.Clear();
        AnsiConsole.MarkupLine("[cyan]██╗    ██╗██████╗  ██████╗ ███████╗     ██████╗██╗     ██╗[/]");
        AnsiConsole.MarkupLine("[cyan]██║    ██║╚════██╗██╔════╝ ██╔════╝    ██╔════╝██║     ██║[/]");
        AnsiConsole.MarkupLine("[cyan]██║ █╗ ██║ █████╔╝███████╗ ███████╗    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine("[cyan]██║███╗██║ ╚═══██╗██╔═══██╗╚════██║    ██║     ██║     ██║[/]");
        AnsiConsole.MarkupLine("[cyan]╚███╔███╔╝██████╔╝╚██████╔╝███████║    ╚██████╗███████╗██║[/]");
        AnsiConsole.MarkupLine("[cyan] ╚══╝╚══╝ ╚═════╝  ╚═════╝ ╚══════╝     ╚═════╝╚══════╝╚═╝[/]");
        AnsiConsole.MarkupLine("[grey]W365 CLI Native v0.1.0 | Bradley Wyatt[/]");
        AnsiConsole.WriteLine();

        if (_session.IsConnected)
        {
            var tenantText = _session.TenantName is not null
                ? $"{_session.TenantName} ({_session.TenantId})"
                : _session.TenantId ?? "unknown";
            AnsiConsole.MarkupLine($"[green]Connected[/] Tenant: [grey]{Markup.Escape(tenantText)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[yellow]Not connected[/]");
        }

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
                .Title("[cyan]Connection[/]")
                .AddChoices(choices));

        switch (choice)
        {
            case "Connect":
                await _session.ConnectAsync();
                Pause();
                break;
            case "Disconnect":
                await _session.DisconnectAsync();
                AnsiConsole.MarkupLine("[green]Disconnected.[/]");
                Pause();
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

        while (true)
        {
            var visibleCloudPcs = FilterCloudPcs(cloudPcs, filter);
            if (selectedIndex >= visibleCloudPcs.Count)
            {
                selectedIndex = Math.Max(0, visibleCloudPcs.Count - 1);
            }

            RenderCloudPcBrowser(cloudPcs, visibleCloudPcs, selectedIndex, filter);
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
                        ShowCloudPcDetails(visibleCloudPcs[selectedIndex]);
                    }
                    break;
                case ConsoleKey.A:
                    if (visibleCloudPcs.Count > 0)
                    {
                        ShowCloudPcDetails(visibleCloudPcs[selectedIndex]);
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

        while (true)
        {
            RenderHeader();
            var cloudAppWidths = GetCloudAppWidths();
            var app = SelectFromTable(
                title: "Windows 365 Cloud Apps",
                header: Row("Status", cloudAppWidths.Status, "Name", cloudAppWidths.Name, "Publisher", cloudAppWidths.Publisher, "Published", cloudAppWidths.Published, "Added", cloudAppWidths.Added),
                items: apps,
                rowFactory: cloudApp => Row(
                    cloudApp.AppStatus ?? "-", cloudAppWidths.Status,
                    cloudApp.DisplayName, cloudAppWidths.Name,
                    cloudApp.Publisher ?? "-", cloudAppWidths.Publisher,
                    cloudApp.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-", cloudAppWidths.Published,
                    cloudApp.AddedDateTime?.ToLocalTime().ToString("g") ?? "-", cloudAppWidths.Added));

            if (app is null)
            {
                return;
            }

            ShowCloudAppDetails(app);
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
                .Title($"[cyan]{Markup.Escape(title)}[/]\n[grey]{Markup.Escape(header)}[/]\n[grey]{Markup.Escape(new string('-', header.Length))}[/]")
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

    private void RenderCloudPcBrowser(
        IReadOnlyList<CloudPcSummary> allCloudPcs,
        IReadOnlyList<CloudPcSummary> visibleCloudPcs,
        int selectedIndex,
        string filter)
    {
        AnsiConsole.Clear();

        var selectedCloudPc = visibleCloudPcs.Count > 0 ? visibleCloudPcs[selectedIndex] : null;
        var grid = new Grid();
        grid.AddColumn();
        grid.AddColumn();
        grid.AddRow(CreateCloudPcTable(allCloudPcs, visibleCloudPcs, selectedIndex, filter), CreateCloudPcSidePanel(selectedCloudPc));

        RenderCompactHeader();
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
        AnsiConsole.MarkupLine("[grey]Up/Down move | PgUp/PgDn page | Enter details | A actions | / filter | C clear | R refresh | Esc back[/]");
    }

    private void RenderCompactHeader()
    {
        if (_session.IsConnected)
        {
            var tenantText = _session.TenantName is not null
                ? $"{_session.TenantName} ({_session.TenantId})"
                : _session.TenantId ?? "unknown";
            AnsiConsole.MarkupLine($"[cyan]W365 CLI Native[/] [grey]v0.1.0 | Bradley Wyatt[/]   [green]Connected[/] [grey]{Markup.Escape(tenantText)}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[cyan]W365 CLI Native[/] [grey]v0.1.0 | Bradley Wyatt[/]   [yellow]Not connected[/]");
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
                selected ? "[white on blue]>[/]" : " ",
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
            new Markup($"[bold]Name[/]\n{Markup.Escape(cloudPc.Name)}"),
            new Markup($"[bold]Status[/] {StatusMarkup(cloudPc.Status)}"),
            new Markup($"[bold]Type[/] {Markup.Escape(cloudPc.ProvisioningType ?? "-")}"),
            new Markup($"[bold]User[/]\n{Markup.Escape(cloudPc.UserPrincipalName ?? "-")}"),
            new Markup($"[bold]Service plan[/]\n{Markup.Escape(cloudPc.ServicePlanName ?? "-")}"),
            new Markup($"[bold]Cloud PC ID[/]\n[grey]{Markup.Escape(cloudPc.Id)}[/]"),
            new Markup("[bold]Actions[/]\n[grey]Enter details, A actions[/]"));

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
                Contains(pc.UserPrincipalName, filter) ||
                Contains(pc.ServicePlanName, filter))
            .ToArray();
    }

    private static bool Contains(string? value, string filter)
    {
        return value?.Contains(filter, StringComparison.OrdinalIgnoreCase) == true;
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

    private static string Selected(string escapedText)
    {
        return $"[white on blue]{escapedText}[/]";
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
        var available = Math.Max(90, Console.WindowWidth - 4);
        const int status = 12;
        const int published = 18;
        const int added = 18;
        var remaining = Math.Max(40, available - status - published - added - 4);
        var name = Math.Max(30, (int)(remaining * 0.65));
        var publisher = Math.Max(18, remaining - name);
        return (status, name, publisher, published, added);
    }

    private static void ShowCloudPcDetails(CloudPcSummary cloudPc)
    {
        var actions = new[]
        {
            "Snapshots",
            "Resize",
            "Restart",
            "Sync",
            "Reprovision",
            "Remote action history",
            "Back"
        };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderCloudPcDetailLayout(cloudPc, actions, selectedActionIndex);
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

                    ShowNotImplementedAction(action);
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

    private static void RenderCloudPcDetailLayout(CloudPcSummary cloudPc, string[] actions, int selectedActionIndex)
    {
        AnsiConsole.MarkupLine($"[cyan]W365 CLI Native > Cloud PCs > {Markup.Escape(cloudPc.Name)}[/]");
        AnsiConsole.WriteLine();

        var details = new Panel(
            new Rows(
                new Markup($"[bold]Name:[/] {Markup.Escape(cloudPc.Name)}"),
                new Markup($"[bold]Status:[/] {StatusMarkup(cloudPc.Status)}"),
                new Markup($"[bold]Type:[/] {Markup.Escape(cloudPc.ProvisioningType ?? "-")}"),
                new Markup($"[bold]User:[/] {Markup.Escape(cloudPc.UserPrincipalName ?? "-")}"),
                new Markup($"[bold]Service plan:[/] {Markup.Escape(cloudPc.ServicePlanName ?? "-")}"),
                new Markup($"[bold]Cloud PC ID:[/] [grey]{Markup.Escape(cloudPc.Id)}[/]")))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions
            .Select((action, index) => index == selectedActionIndex
                ? $"[white on blue]> {Markup.Escape(action)}[/]"
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
    }

    private static void ShowNotImplementedAction(string action)
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(action)} is not implemented in the native CLI yet.[/]");
        AnsiConsole.MarkupLine("[grey]The action shell is in place. The next native milestone is wiring each action to Graph.[/]");
        Pause();
    }

    private static void ShowCloudAppDetails(CloudAppSummary app)
    {
        AnsiConsole.Clear();
        var panel = new Panel(
            new Rows(
                new Markup($"[bold]Name:[/] {Markup.Escape(app.DisplayName)}"),
                new Markup($"[bold]Status:[/] {Markup.Escape(app.AppStatus ?? "-")}"),
                new Markup($"[bold]Publisher:[/] {Markup.Escape(app.Publisher ?? "-")}"),
                new Markup($"[bold]Discovered app:[/] {Markup.Escape(app.DiscoveredAppName ?? "-")}"),
                new Markup($"[bold]Added:[/] {Markup.Escape(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
                new Markup($"[bold]Published:[/] {Markup.Escape(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-")}"),
                new Markup($"[bold]Cloud app ID:[/] {Markup.Escape(app.Id)}")))
            .Header("Cloud App details")
            .Border(BoxBorder.Rounded);

        AnsiConsole.Write(panel);
        Pause();
    }

    private sealed record TableChoice<T>(string Label, T? Item, bool IsBack);

    private sealed record MenuChoice(string Key, string Category, string Title, string Description, string Accent);

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
        AnsiConsole.MarkupLine("[cyan]W365 CLI Native[/]");
        AnsiConsole.MarkupLine("A native .NET rewrite experiment for Windows 365 Cloud PC workflows.");
        AnsiConsole.MarkupLine("This project does not depend on the PowerShell W365CLI module.");
        Pause();
    }

    private static void Pause()
    {
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to continue...[/]");
        Console.ReadKey(intercept: true);
    }
}
