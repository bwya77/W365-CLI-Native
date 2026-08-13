using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

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
            var choice = PromptChoice(
                () => { },
                "[#58a6ff]Reports[/]",
                ["Sign-in status", "Connectivity history", "Launch details", "Cloud PC reports", "Back"],
                "Back");

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
}
