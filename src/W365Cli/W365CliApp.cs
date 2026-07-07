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

            var choice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("[cyan]Choose an area[/]")
                    .PageSize(10)
                    .AddChoices(
                        "Connection - connect or disconnect Graph",
                        "Cloud PCs - browse inventory",
                        "Cloud Apps - browse, publish, unpublish",
                        "About",
                        "Exit"));

            switch (choice)
            {
                case "Connection - connect or disconnect Graph":
                    await ShowConnectionAsync();
                    break;
                case "Cloud PCs - browse inventory":
                    await ShowCloudPcsAsync();
                    break;
                case "Cloud Apps - browse, publish, unpublish":
                    await ShowCloudAppsAsync();
                    break;
                case "About":
                    ShowAbout();
                    break;
                case "Exit":
                    return 0;
            }
        }
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
            ? new[] { "Disconnect", "Reconnect", "Back" }
            : new[] { "Connect", "Back" };

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("[cyan]Connection[/]")
                .AddChoices(choices));

        switch (choice)
        {
            case "Connect":
            case "Reconnect":
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

    private async Task ShowCloudPcsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var cloudPcs = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud PCs...", async _ => await _session.Graph.GetCloudPcsAsync());

        if (cloudPcs.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No Cloud PCs returned.[/]");
            Pause();
            return;
        }

        while (true)
        {
            RenderHeader();
            var cloudPcWidths = GetCloudPcWidths();
            var cloudPc = SelectFromTable(
                title: "Windows 365 Cloud PCs",
                header: Row("Name", cloudPcWidths.Name, "Status", cloudPcWidths.Status, "Type", cloudPcWidths.Type, "User", cloudPcWidths.User, "Service plan", cloudPcWidths.ServicePlan),
                items: cloudPcs,
                rowFactory: pc => Row(
                    pc.Name, cloudPcWidths.Name,
                    pc.Status ?? "-", cloudPcWidths.Status,
                    pc.ProvisioningType ?? "-", cloudPcWidths.Type,
                    pc.UserPrincipalName ?? "-", cloudPcWidths.User,
                    pc.ServicePlanName ?? "-", cloudPcWidths.ServicePlan));

            if (cloudPc is null)
            {
                return;
            }

            ShowCloudPcDetails(cloudPc);
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
        var name = Math.Max(24, (int)(remaining * 0.32));
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
        AnsiConsole.Clear();
        var panel = new Panel(
            new Rows(
                new Markup($"[bold]Name:[/] {Markup.Escape(cloudPc.Name)}"),
                new Markup($"[bold]Status:[/] {Markup.Escape(cloudPc.Status ?? "-")}"),
                new Markup($"[bold]Type:[/] {Markup.Escape(cloudPc.ProvisioningType ?? "-")}"),
                new Markup($"[bold]User:[/] {Markup.Escape(cloudPc.UserPrincipalName ?? "-")}"),
                new Markup($"[bold]Service plan:[/] {Markup.Escape(cloudPc.ServicePlanName ?? "-")}"),
                new Markup($"[bold]Cloud PC ID:[/] {Markup.Escape(cloudPc.Id)}")))
            .Header("Cloud PC details")
            .Border(BoxBorder.Rounded);

        AnsiConsole.Write(panel);
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
