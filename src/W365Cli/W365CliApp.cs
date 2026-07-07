using Spectre.Console;

namespace W365Cli;

internal sealed class W365CliApp
{
    private readonly W365Session _session = new();

    public async Task<int> RunAsync(string[] args)
    {
        Console.Title = "W365 CLI Native";

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
            AnsiConsole.MarkupLine($"[green]Connected[/] Tenant: [grey]{_session.TenantId ?? "unknown"}[/]");
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
                _session.Disconnect();
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

        var table = new Table()
            .Title("Windows 365 Cloud PCs")
            .AddColumn("Name")
            .AddColumn("Status")
            .AddColumn("Type")
            .AddColumn("User")
            .AddColumn("Service plan");

        foreach (var pc in cloudPcs)
        {
            table.AddRow(
                Markup.Escape(pc.Name),
                Markup.Escape(pc.Status ?? "-"),
                Markup.Escape(pc.ProvisioningType ?? "-"),
                Markup.Escape(pc.UserPrincipalName ?? "-"),
                Markup.Escape(pc.ServicePlanName ?? "-"));
        }

        AnsiConsole.Write(table);
        Pause();
    }

    private async Task ShowCloudAppsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var apps = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading Cloud Apps...", async _ => await _session.Graph.GetCloudAppsAsync());

        if (apps.Count == 0)
        {
            AnsiConsole.MarkupLine("[yellow]No Cloud Apps returned.[/]");
            Pause();
            return;
        }

        var table = new Table()
            .Title("Windows 365 Cloud Apps")
            .AddColumn("Status")
            .AddColumn("Name")
            .AddColumn("Publisher")
            .AddColumn("Added")
            .AddColumn("Published");

        foreach (var app in apps)
        {
            table.AddRow(
                Markup.Escape(app.AppStatus ?? "-"),
                Markup.Escape(app.DisplayName),
                Markup.Escape(app.Publisher ?? "-"),
                Markup.Escape(app.AddedDateTime?.ToLocalTime().ToString("g") ?? "-"),
                Markup.Escape(app.LastPublishedDateTime?.ToLocalTime().ToString("g") ?? "-"));
        }

        AnsiConsole.Write(table);
        Pause();
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
