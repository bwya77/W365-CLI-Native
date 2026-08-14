using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

    private async Task ShowProvisioningAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var policies = await LoadProvisioningPoliciesAsync();
        if (policies.Count == 0)
        {
            TimedMessage("[yellow]No provisioning policies were returned. Use \"Create policy\" to add one.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = ProvisioningPolicySortMode.Name;
        while (true)
        {
            var visiblePolicies = SortProvisioningPolicies(FilterProvisioningPolicies(policies, filter), sortMode);
            if (visiblePolicies.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visiblePolicies.Count)
            {
                selectedIndex = visiblePolicies.Count - 1;
            }

            AnsiConsole.Clear();
            RenderProvisioningPolicyBrowser(policies, visiblePolicies, selectedIndex, filter, sortMode);
            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, visiblePolicies.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.PageUp:
                    selectedIndex = Math.Max(0, selectedIndex - 10);
                    break;
                case ConsoleKey.PageDown:
                    selectedIndex = Math.Min(Math.Max(0, visiblePolicies.Count - 1), selectedIndex + 10);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, visiblePolicies.Count - 1);
                    break;
                case ConsoleKey.Enter:
                case ConsoleKey.A:
                    if (visiblePolicies.Count > 0)
                    {
                        await ShowProvisioningPolicyDetailsAsync(visiblePolicies[selectedIndex]);
                        policies = await LoadProvisioningPoliciesAsync();
                    }
                    break;
                case ConsoleKey.R:
                    policies = await LoadProvisioningPoliciesAsync();
                    selectedIndex = 0;
                    break;
                case ConsoleKey.C:
                    filter = string.Empty;
                    selectedIndex = 0;
                    break;
                case ConsoleKey.S:
                    sortMode = NextProvisioningPolicySortMode(sortMode);
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

    private async Task<IReadOnlyList<ProvisioningPolicySummary>> LoadProvisioningPoliciesAsync()
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading provisioning policies...", async _ => await _session.Graph.GetProvisioningPoliciesAsync());
    }

    private void RenderProvisioningPolicyBrowser(
        IReadOnlyList<ProvisioningPolicySummary> allPolicies,
        IReadOnlyList<ProvisioningPolicySummary> visiblePolicies,
        int selectedIndex,
        string filter,
        ProvisioningPolicySortMode sortMode)
    {
        RenderBreadcrumb("Provisioning", "Policies");
        AnsiConsole.Write(CreateProvisioningPolicySummaryPanel(allPolicies, visiblePolicies, filter));
        AnsiConsole.Write(CreateProvisioningPolicyTable(visiblePolicies, selectedIndex));
        AnsiConsole.MarkupLine($"[grey]Sort: {FormatProvisioningPolicySortMode(sortMode)} | Up/Down move | Enter actions | / filter | C clear | S sort | R refresh | Esc/B/Q back | P or Ctrl+K command palette[/]");
        RenderStatusBar();
    }

    private static Panel CreateProvisioningPolicySummaryPanel(IReadOnlyList<ProvisioningPolicySummary> allPolicies, IReadOnlyList<ProvisioningPolicySummary> visiblePolicies, string filter)
    {
        var typeSummary = string.Join("  ", allPolicies
            .GroupBy(policy => policy.ProvisioningType ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));
        var joinSummary = string.Join("  ", allPolicies
            .GroupBy(policy => policy.DomainJoinTypes ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}"));

        return new Panel(new Rows(
                new Markup($"[white]Total[/] {allPolicies.Count}   [white]Visible[/] {visiblePolicies.Count}   [white]Filter[/] {Markup.Escape(string.IsNullOrWhiteSpace(filter) ? "none" : filter)}"),
                new Markup($"[white]Types[/] {Markup.Escape(typeSummary)}"),
                new Markup($"[white]Join[/] {Markup.Escape(joinSummary)}")))
            .Header("Policies")
            .Border(BoxBorder.Rounded);
    }

    private static Table CreateProvisioningPolicyTable(IReadOnlyList<ProvisioningPolicySummary> visiblePolicies, int selectedIndex)
    {
        var widths = GetProvisioningPolicyWidths();
        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn(" ")
            .AddColumn("Name")
            .AddColumn("Type")
            .AddColumn("Image")
            .AddColumn("Join")
            .AddColumn("SSO");

        var showGroups = Console.WindowWidth >= 130;
        if (showGroups)
        {
            table.AddColumn("Groups");
        }

        if (visiblePolicies.Count == 0)
        {
            var cells = new List<string> { "-", "[grey]No policies match the current filter.[/]", "-", "-", "-", "-" };
            if (showGroups) { cells.Add("-"); }
            table.AddRow(cells.ToArray());
            return table;
        }

        var pageSize = Math.Max(8, Math.Min(18, Console.WindowHeight - 16));
        var start = Math.Clamp(selectedIndex - pageSize / 2, 0, Math.Max(0, visiblePolicies.Count - pageSize));
        var visible = visiblePolicies.Skip(start).Take(pageSize).ToArray();
        for (var index = 0; index < visible.Length; index++)
        {
            var absoluteIndex = start + index;
            var policy = visible[index];
            var selected = absoluteIndex == selectedIndex;
            var row = new List<string>
            {
                selected ? "[black on #58a6ff]>[/]" : " ",
                selected ? Selected(Markup.Escape(Fit(policy.DisplayName, widths.Name))) : Markup.Escape(Fit(policy.DisplayName, widths.Name)),
                selected ? Selected(Markup.Escape(Fit(policy.ProvisioningType ?? "-", widths.Type))) : Markup.Escape(Fit(policy.ProvisioningType ?? "-", widths.Type)),
                selected ? Selected(Markup.Escape(Fit(policy.ImageDisplayName ?? "-", widths.Image))) : Markup.Escape(Fit(policy.ImageDisplayName ?? "-", widths.Image)),
                selected ? Selected(Markup.Escape(Fit(policy.DomainJoinTypes ?? "-", widths.Join))) : Markup.Escape(Fit(policy.DomainJoinTypes ?? "-", widths.Join)),
                selected ? Selected(Markup.Escape(Fit(FormatBool(policy.EnableSingleSignOn), widths.Sso))) : Markup.Escape(Fit(FormatBool(policy.EnableSingleSignOn), widths.Sso))
            };
            if (showGroups)
            {
                row.Add(selected ? Selected(Markup.Escape(Fit(string.Join(", ", policy.AssignedGroupNames), widths.Groups))) : Markup.Escape(Fit(string.Join(", ", policy.AssignedGroupNames), widths.Groups)));
            }

            table.AddRow(row.ToArray());
        }

        return table;
    }

    private static (int Name, int Type, int Image, int Join, int Sso, int Groups) GetProvisioningPolicyWidths()
    {
        var available = Math.Max(95, Console.WindowWidth - 4);
        const int type = 12;
        const int join = 16;
        const int sso = 5;
        var showGroups = Console.WindowWidth >= 130;
        var groups = showGroups ? Math.Max(20, (int)(available * 0.20)) : 0;
        var remaining = Math.Max(45, available - type - join - sso - groups - (showGroups ? 6 : 5));
        var name = Math.Clamp((int)(remaining * 0.55), 28, 44);
        var image = Math.Max(22, remaining - name);
        return (name, type, image, join, sso, groups);
    }

    private async Task ShowProvisioningPolicyDetailsAsync(ProvisioningPolicySummary policy)
    {
        var actions = GetProvisioningPolicyActions(policy);
        var selectedActionIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderProvisioningPolicyDetailLayout(policy, actions, selectedActionIndex);
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
                    await InvokeProvisioningPolicyActionAsync(policy, action);
                    if (action is "Delete")
                    {
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
                    if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
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

    private static string[] GetProvisioningPolicyActions(ProvisioningPolicySummary policy)
    {
        var actions = new List<string> { "View Cloud PCs", "Export", "Create copy", "Reprovision policy Cloud PCs" };

        if (IsSharedProvisioningPolicy(policy))
        {
            actions.Add("Reprovision (reserve %)");
            actions.Add("Check reprovision status");
        }

        if (IsSharedByEntraGroupPolicy(policy))
        {
            actions.Add("User experience sync");
        }

        if (policy.AssignedGroupIds.Count > 0)
        {
            actions.Add("Manage group members");
        }

        actions.Add("Delete");
        actions.Add("Back");
        return actions.ToArray();
    }

    private static bool IsSharedProvisioningPolicy(ProvisioningPolicySummary policy)
    {
        return policy.ProvisioningType is not null &&
            new[] { "shared", "sharedByUser", "sharedByEntraGroup" }.Any(value =>
                string.Equals(policy.ProvisioningType, value, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// User settings persistence ("user experience sync") is only available for shared-by-Entra-
    /// group policies — not dedicated or shared-by-user — per Microsoft's documentation.
    /// </summary>
    private static bool IsSharedByEntraGroupPolicy(ProvisioningPolicySummary policy)
    {
        return string.Equals(policy.ProvisioningType, "sharedByEntraGroup", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderProvisioningPolicyDetailLayout(ProvisioningPolicySummary policy, IReadOnlyList<string> actions, int selectedActionIndex)
    {
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName);
        var details = new Panel(new Rows(
                new Markup(PropertyInline("Name", policy.DisplayName)),
                new Markup(PropertyInline("Description", policy.Description ?? "-")),
                new Markup(PropertyInline("Type", policy.ProvisioningType ?? "-")),
                new Markup(PropertyInline("Image", policy.ImageDisplayName ?? "-")),
                new Markup(PropertyInline("Image type", policy.ImageType ?? "-")),
                new Markup(PropertyInline("Domain join", policy.DomainJoinTypes ?? "-")),
                new Markup(PropertyInline("Single sign-on", FormatBool(policy.EnableSingleSignOn))),
                new Markup(PropertyInline("Local admin", FormatBool(policy.LocalAdminEnabled))),
                new Markup(PropertyInline("Naming template", policy.CloudPcNamingTemplate ?? "-")),
                new Markup(PropertyInline("Cloud PC group", policy.CloudPcGroupDisplayName ?? "-")),
                new Markup(PropertyInline("Managed by", policy.ManagedBy ?? "-")),
                new Markup(PropertyInline("Grace period hours", policy.GracePeriodInHours?.ToString() ?? "-")),
                new Markup(PropertyBlock("Assigned groups", string.Join(", ", policy.AssignedGroupNames))),
                new Markup(PropertyBlock("Policy ID", policy.Id))))
            .Header("Details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
        var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
            .Header("Actions")
            .Border(BoxBorder.Rounded);

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, actionPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(actionPanel);
        }

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back | P or Ctrl+K command palette[/]");
    }

    private async Task InvokeProvisioningPolicyActionAsync(ProvisioningPolicySummary policy, string action)
    {
        switch (action)
        {
            case "View Cloud PCs":
                await ShowCloudPcsForProvisioningPolicyAsync(policy);
                break;
            case "Export":
                await ExportProvisioningPolicyAsync(policy);
                break;
            case "Create copy":
                await CreateProvisioningPolicyCopyAsync(policy);
                break;
            case "Reprovision policy Cloud PCs":
                await ReprovisionProvisioningPolicyCloudPcsAsync(policy);
                break;
            case "Reprovision (reserve %)":
                await ApplyProvisioningPolicyReservePercentageAsync(policy);
                break;
            case "Check reprovision status":
                await ShowProvisioningPolicyApplyStatusAsync(policy);
                break;
            case "User experience sync":
                await ShowUserExperienceSyncAsync(policy);
                break;
            case "Manage group members":
                await ShowProvisioningPolicyGroupMembersAsync(policy);
                break;
            case "Delete":
                await DeleteProvisioningPolicyWithGuardAsync(policy);
                break;
        }
    }

    /// <summary>
    /// Windows 365 requires a provisioning policy to have zero group assignments before it can be
    /// deleted — attempting to delete an assigned policy fails with a 400 Bad Request. Detect that
    /// up front and offer to remove the assignments (via the assign action with an empty list)
    /// before retrying the delete, instead of surfacing a confusing raw error.
    /// </summary>
    private async Task DeleteProvisioningPolicyWithGuardAsync(ProvisioningPolicySummary policy)
    {
        if (policy.AssignedGroupNames.Count > 0)
        {
            void RenderContext()
            {
                RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Delete");
                AnsiConsole.MarkupLine("[yellow]This policy still has group assignments.[/]");
                AnsiConsole.MarkupLine("[grey]Windows 365 requires a provisioning policy to have no assignments before it can be deleted.[/]");
                AnsiConsole.MarkupLine($"[grey]Assigned groups:[/] {Markup.Escape(string.Join(", ", policy.AssignedGroupNames))}");
                AnsiConsole.WriteLine();
            }

            var choice = PromptChoice(RenderContext, "How would you like to proceed?", ["Remove assignments, then delete", "Cancel"], "Cancel");

            if (choice != "Remove assignments, then delete")
            {
                TimedMessage("[yellow]Delete cancelled.[/]");
                return;
            }

            try
            {
                await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Removing assignments...", async _ => await _session.Graph.UnassignProvisioningPolicyAsync(policy.Id));
            }
            catch (Exception ex)
            {
                if (await HandlePermissionErrorAsync(ex, "Remove assignments", policy.DisplayName) ||
                    HandleLockedResourceError(ex, "Remove assignments", policy.DisplayName))
                {
                    return;
                }

                ShowActionResult("Failed", "Remove assignments", policy.DisplayName, "[red]Failed to remove assignments.[/]", ex.Message);
                return;
            }
        }

        await ConfirmAndRunAsync("Delete", policy.DisplayName, async () => await _session.Graph.DeleteProvisioningPolicyAsync(policy.Id), "Policy", policy.DisplayName);
    }

    private async Task ShowCloudPcsForProvisioningPolicyAsync(ProvisioningPolicySummary policy)
    {
        var cloudPcs = await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync("Loading policy Cloud PCs...", async _ => await _session.Graph.GetCloudPcsByProvisioningPolicyAsync(policy.Id));
        if (cloudPcs.Count == 0)
        {
            TimedMessage("[yellow]No Cloud PCs were returned for this provisioning policy.[/]");
            return;
        }

        var selectedIndex = 0;
        var filter = string.Empty;
        var sortMode = CloudPcSortMode.Name;
        while (true)
        {
            var visibleCloudPcs = SortCloudPcs(FilterCloudPcs(cloudPcs, filter), sortMode);
            if (visibleCloudPcs.Count == 0)
            {
                selectedIndex = 0;
            }
            else if (selectedIndex >= visibleCloudPcs.Count)
            {
                selectedIndex = visibleCloudPcs.Count - 1;
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Cloud PCs");
            AnsiConsole.Write(CreateCloudPcSummaryPanel(cloudPcs, visibleCloudPcs, filter));
            AnsiConsole.Write(CreateCloudPcTable(cloudPcs, visibleCloudPcs, selectedIndex, filter));
            AnsiConsole.MarkupLine($"[grey]Sort: {FormatCloudPcSortMode(sortMode)} | Enter actions | D disk | N snapshots | Z resize | Y sync | / filter | C clear | S sort | R refresh | Esc/B/Q back[/]");
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
                    cloudPcs = await _session.Graph.GetCloudPcsByProvisioningPolicyAsync(policy.Id);
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
                default:
                    if (key.KeyChar is '/' or 'f' or 'F')
                    {
                        filter = PromptFilter();
                        selectedIndex = 0;
                    }
                    else if (key.KeyChar is 'b' or 'B' or 'q' or 'Q')
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
                    break;
            }
        }
    }

    private async Task ExportProvisioningPolicyAsync(ProvisioningPolicySummary policy)
    {
        var defaultPath = Path.Combine(Environment.CurrentDirectory, $"{SanitizeFileName(policy.DisplayName)}.json");
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Export");
        var path = AnsiConsole.Prompt(
            new TextPrompt<string>($"Export path [[{Markup.Escape(defaultPath)}]]:")
                .AllowEmpty());
        if (string.IsNullOrWhiteSpace(path))
        {
            path = defaultPath;
        }

        var exportJson = _session.Graph.ExportProvisioningPolicyJson(policy);
        await File.WriteAllTextAsync(path, exportJson);
        ShowActionResult("Exported", "Export", path, "[green]Exported.[/]");
    }

    private async Task CreateProvisioningPolicyCopyAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Create copy");
        var displayName = AnsiConsole.Ask<string>("New policy display name:");
        if (string.IsNullOrWhiteSpace(displayName))
        {
            TimedMessage("[yellow]Create copy cancelled. Display name is required.[/]");
            return;
        }

        var assign = AskYesNo("Recreate assignment targets on the new policy?");
        await ConfirmAndRunAsync(
            "Create copy",
            $"{policy.DisplayName} to {displayName}",
            async () => await _session.Graph.CreateProvisioningPolicyCopyAsync(policy, displayName, assign),
            "Policy",
            policy.DisplayName);
    }

    private async Task CreateProvisioningPolicyWizardAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Create policy");
        AnsiConsole.MarkupLine("[grey]Create a new Windows 365 provisioning policy.[/]");
        AnsiConsole.WriteLine();

        var displayName = AnsiConsole.Ask<string>("Policy display name:");
        if (string.IsNullOrWhiteSpace(displayName))
        {
            TimedMessage("[yellow]Create policy cancelled. Display name is required.[/]");
            return;
        }

        var description = AnsiConsole.Prompt(new TextPrompt<string>("Description [[optional]]:").AllowEmpty());

        void RenderPolicyWizardContext()
        {
            RenderBreadcrumb("Provisioning", "Create policy");
            AnsiConsole.MarkupLine($"[grey]Name:[/] {Markup.Escape(displayName)}");
            AnsiConsole.WriteLine();
        }

        // Mirrors the Windows 365 admin portal's own wizard structure (License type, then Flex
        // mode, then Experience) instead of exposing raw Graph provisioningType enum names —
        // "Dedicated"/"Shared by user"/"Shared by Entra group" reads as three unrelated options
        // when it's really "which license, and if Flex, which mode."
        var licenseTypeChoice = PromptChoice(
            RenderPolicyWizardContext,
            "License type",
            ["Enterprise", "Windows 365 Flex", "Reserve", "Back"],
            "Back");
        if (licenseTypeChoice == "Back")
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        string provisioningType;
        string? flexModeChoice = null;
        if (licenseTypeChoice == "Windows 365 Flex")
        {
            flexModeChoice = PromptChoice(
                RenderPolicyWizardContext,
                "Windows 365 Flex mode",
                ["Dedicated", "Shared", "Back"],
                "Back");
            if (flexModeChoice == "Back")
            {
                TimedMessage("[yellow]Create policy cancelled.[/]");
                return;
            }

            // Flex Dedicated is still Graph's plain "dedicated" provisioningType -- the Enterprise
            // vs. Flex distinction is entirely about which purchased license SKU the assigned
            // group draws from, which Graph resolves automatically; it isn't a separate value in
            // the policy body. Flex Shared is what Graph calls "sharedByEntraGroup".
            provisioningType = flexModeChoice == "Shared" ? "sharedByEntraGroup" : "dedicated";
        }
        else if (licenseTypeChoice == "Reserve")
        {
            provisioningType = "reserve";
        }
        else
        {
            provisioningType = "dedicated";
        }

        var licenseLabel = flexModeChoice is null ? licenseTypeChoice : $"{licenseTypeChoice} ({flexModeChoice})";

        // "Access only apps" (Cloud Apps) is only valid for Windows 365 Flex in Shared mode —
        // Graph requires provisioningType=sharedByEntraGroup for userExperienceType=cloudApp, so
        // every other combination stays full-desktop without even asking.
        var userExperienceType = "cloudPc";
        if (string.Equals(provisioningType, "sharedByEntraGroup", StringComparison.OrdinalIgnoreCase))
        {
            var experienceChoice = PromptChoice(
                RenderPolicyWizardContext,
                "Experience",
                ["Access a full Cloud PC desktop", "Access only apps (Cloud Apps)", "Back"],
                "Back");
            if (experienceChoice == "Back")
            {
                TimedMessage("[yellow]Create policy cancelled.[/]");
                return;
            }

            userExperienceType = experienceChoice.Contains("Cloud Apps", StringComparison.OrdinalIgnoreCase) ? "cloudApp" : "cloudPc";
        }

        // sharedByEntraGroup requires a purchased Frontline "shared" service plan SKU in the
        // tenant -- without one, Graph rejects policy creation with a generic, unhelpful
        // "400 badRequest: malformed or incorrect" and no further detail. Checking servicePlans'
        // provisioningType up front turns that into a clear, specific warning before the user
        // spends time picking an image/region/naming template only to hit a vague failure at the
        // very end.
        if (!string.Equals(provisioningType, "dedicated", StringComparison.OrdinalIgnoreCase))
        {
            IReadOnlyList<CloudPcServicePlan> servicePlans;
            try
            {
                servicePlans = await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Checking available licenses...", async _ => await _session.Graph.GetCloudPcServicePlansAsync());
            }
            catch
            {
                servicePlans = [];
            }

            var hasMatchingPlan = servicePlans.Any(plan =>
                string.Equals(plan.ProvisioningType, provisioningType, StringComparison.OrdinalIgnoreCase));

            if (servicePlans.Count > 0 && !hasMatchingPlan)
            {
                var proceed = AskYesNo(
                    $"No purchased \"{licenseLabel}\" license was found in this tenant, so Graph will likely reject this with a generic 400 error. Continue anyway?",
                    defaultToYes: false);
                if (!proceed)
                {
                    TimedMessage("[yellow]Create policy cancelled.[/]");
                    return;
                }
            }
        }

        IReadOnlyList<GraphTableRow> images;
        try
        {
            images = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading gallery images...", async _ => await _session.Graph.GetGalleryImageRowsAsync());
        }
        catch (Exception ex)
        {
            if (!await HandlePermissionErrorAsync(ex, "Create policy", displayName) &&
                !HandleLockedResourceError(ex, "Create policy", displayName))
            {
                ShowActionResult("Failed", "Create policy", displayName, "[red]Failed to load gallery images.[/]", ex.Message);
            }
            return;
        }

        if (images.Count == 0)
        {
            TimedMessage("[yellow]No gallery images are available to select.[/]");
            return;
        }

        // notSupported gallery images can't be used to provision new Cloud PCs and Graph rejects
        // the create with a generic 400 if you pick one — filter them out so only images that
        // can actually be selected are shown.
        var selectableImages = images
            .Where(image => !string.Equals(GetOptionalField(image, "status"), "notSupported", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (selectableImages.Length == 0)
        {
            selectableImages = images.ToArray();
        }

        var imageHeader = Row("Image", 50, "Status", 14, "OS version", 16);
        var selectedImage = SelectFromTable(
            "Select gallery image",
            imageHeader,
            selectableImages,
            image => Row(image.Title, 50, GetField(image, "status"), 14, GetField(image, "osVersionNumber"), 16));
        if (selectedImage is null)
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        var imageId = GetOptionalField(selectedImage, "id", "Id", "ID");
        if (string.IsNullOrWhiteSpace(imageId))
        {
            TimedMessage("[red]Selected image is missing an ID; cannot continue.[/]");
            return;
        }

        var namingTemplate = AnsiConsole.Prompt(
            new TextPrompt<string>("Cloud PC naming template:")
                .DefaultValue("CPC-%USERNAME:5%-%RAND:5%"));

        IReadOnlyList<GraphTableRow> regions;
        try
        {
            regions = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading supported regions...", async _ => await _session.Graph.GetSupportedRegionRowsAsync());
        }
        catch (Exception ex)
        {
            if (!await HandlePermissionErrorAsync(ex, "Create policy", displayName) &&
                !HandleLockedResourceError(ex, "Create policy", displayName))
            {
                ShowActionResult("Failed", "Create policy", displayName, "[red]Failed to load supported regions.[/]", ex.Message);
            }
            return;
        }

        var availableRegions = regions
            .Where(region => string.IsNullOrWhiteSpace(GetOptionalField(region, "supportedSolution")) ||
                string.Equals(GetOptionalField(region, "supportedSolution"), "windows365", StringComparison.OrdinalIgnoreCase))
            .Where(region => string.Equals(GetOptionalField(region, "regionStatus"), "available", StringComparison.OrdinalIgnoreCase))
            .GroupBy(region => region.Title, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        if (availableRegions.Length == 0)
        {
            availableRegions = regions
                .GroupBy(region => region.Title, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
        }

        if (availableRegions.Length == 0)
        {
            TimedMessage("[yellow]No supported regions were returned; a region is required for Microsoft Entra join.[/]");
            return;
        }

        // Microsoft Entra join requires either a region or an on-premises network connection —
        // leaving both empty causes Graph to reject the create with a generic 400, so a region
        // selection here is mandatory rather than optional.
        var regionOptions = availableRegions.Select(region => region.Title).Append("Back").ToArray();
        var regionChoice = PromptChoice(
            () =>
            {
                RenderBreadcrumb("Provisioning", "Create policy");
                AnsiConsole.MarkupLine($"[grey]Name:[/] {Markup.Escape(displayName)}");
                AnsiConsole.WriteLine();
            },
            "Region for Microsoft Entra joined Cloud PCs",
            regionOptions,
            "Back");

        if (regionChoice == "Back")
        {
            TimedMessage("[yellow]Create policy cancelled.[/]");
            return;
        }

        var regionRow = availableRegions.First(region => region.Title == regionChoice);
        var regionName = regionRow.Title;

        var enableSso = AskYesNo("Enable single sign-on?", defaultToYes: false);
        var localAdmin = AskYesNo("Enable local admin?", defaultToYes: false);

        var isSharedByEntraGroup = string.Equals(provisioningType, "sharedByEntraGroup", StringComparison.OrdinalIgnoreCase);
        var (assignGroupId, _) = await PromptForEntraGroupAsync(
            required: isSharedByEntraGroup,
            reasonIfRequired: "Windows 365 Flex Shared policies need a group assignment to actually provision any Cloud PCs.");

        await ConfirmAndRunAsync(
            "Create policy",
            displayName,
            async () => await _session.Graph.CreateProvisioningPolicyAsync(
                displayName,
                description,
                provisioningType,
                imageId,
                selectedImage.Title,
                "gallery",
                "azureADJoin",
                regionName,
                namingTemplate,
                enableSso,
                localAdmin,
                assignGroupId,
                userExperienceType),
            "Policy",
            displayName);
    }

    /// <summary>
    /// Search-and-pick flow for choosing an Entra group, replacing a raw object-ID paste box that
    /// gave users no way to actually find the group they wanted. Loops on empty results or "Back"
    /// from the picker (returning to the search prompt, not aborting the whole wizard) until a
    /// group is chosen or the user explicitly skips (only allowed when not required).
    /// </summary>
    private async Task<(string? Id, string? Name)> PromptForEntraGroupAsync(bool required, string reasonIfRequired)
    {
        while (true)
        {
            var searchTerm = AnsiConsole.Prompt(
                new TextPrompt<string>(required
                    ? "Search for the Entra group to assign this policy to:"
                    : "Search for an Entra group to assign now [[optional — leave blank to skip]]:")
                    .AllowEmpty());

            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                if (!required)
                {
                    return (null, null);
                }

                var skipAnyway = AskYesNo($"{reasonIfRequired} Continue without assigning a group now?", defaultToYes: false);
                if (skipAnyway)
                {
                    return (null, null);
                }

                continue;
            }

            IReadOnlyList<EntraGroupSummary> groups;
            try
            {
                groups = await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync("Searching groups...", async _ => await _session.Graph.SearchGroupsAsync(searchTerm.Trim()));
            }
            catch (Exception ex)
            {
                ShowActionResult("Failed", "Search groups", searchTerm, "[red]Failed to search Entra groups.[/]", ex.Message);
                continue;
            }

            if (groups.Count == 0)
            {
                TimedMessage("[yellow]No matching groups found. Try a different search term.[/]");
                continue;
            }

            var groupHeader = Row("Group", 50, "Mail nickname", 30);
            var selectedGroup = SelectFromTable(
                "Select Entra group",
                groupHeader,
                groups,
                group => Row(group.Name, 50, group.MailNickname ?? "-", 30));

            if (selectedGroup is not null)
            {
                return (selectedGroup.Id, selectedGroup.Name);
            }
        }
    }

    private async Task ReprovisionProvisioningPolicyCloudPcsAsync(ProvisioningPolicySummary policy)
    {
        void RenderContext() => RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Reprovision");

        var osChoice = PromptChoice(RenderContext, "Policy reprovision OS version", ["Keep policy/default", "Windows 11", "Windows 10", "Back"], "Back");
        if (osChoice == "Back")
        {
            return;
        }

        var accountChoice = PromptChoice(RenderContext, "Policy reprovision user account type", ["Keep policy/default", "Standard user", "Administrator", "Back"], "Back");
        if (accountChoice == "Back")
        {
            return;
        }

        RenderContext();

        AnsiConsole.WriteLine();
        var excludeInput = AnsiConsole.Prompt(
            new TextPrompt<string>("Exclude Cloud PCs by name, ID, or UPN, comma-separated [[optional]]:")
                .AllowEmpty());
        var exclusions = string.IsNullOrWhiteSpace(excludeInput)
            ? []
            : excludeInput.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var osVersion = osChoice switch
        {
            "Windows 11" => "windows11",
            "Windows 10" => "windows10",
            _ => null
        };
        var accountType = accountChoice switch
        {
            "Standard user" => "standardUser",
            "Administrator" => "administrator",
            _ => null
        };

        await ConfirmAndRunAsync(
            "Reprovision",
            $"{policy.DisplayName} policy Cloud PCs",
            async () => await _session.Graph.ReprovisionCloudPcsByPolicyAsync(policy.Id, osVersion, accountType, exclusions),
            "Policy",
            policy.DisplayName);
    }

    private async Task ApplyProvisioningPolicyReservePercentageAsync(ProvisioningPolicySummary policy)
    {
        var reservePercentage = PromptForInteger(
            $"Reprovision (reserve %) — {policy.DisplayName}",
            "Reprovisions Frontline shared Cloud PCs under this policy while keeping the given percentage available. Cloud PCs are only reprovisioned once they have no active connected user. Enter the percentage to keep available (0-99), or Esc/B/Q to cancel.",
            0,
            99);
        if (reservePercentage is null)
        {
            TimedMessage("[yellow]Reprovision cancelled.[/]");
            return;
        }

        var forceLogoff = reservePercentage == 0 &&
            AskYesNo("Forcibly sign out connected users and reprovision immediately?", defaultToYes: false);

        await ConfirmAndRunAsync(
            "Reprovision (reserve %)",
            $"{policy.DisplayName} — reserve {reservePercentage}%{(forceLogoff ? ", force logoff" : string.Empty)}",
            async () => await _session.Graph.ApplyProvisioningPolicyAsync(policy.Id, reservePercentage.Value, forceLogoff),
            "Policy",
            policy.DisplayName);
    }

    private async Task ShowProvisioningPolicyApplyStatusAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Reprovision status");

        CloudPcPolicyApplyActionResult? result;
        try
        {
            result = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Checking reprovision status...", async _ => await _session.Graph.GetProvisioningPolicyApplyActionResultAsync(policy.Id));
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Check reprovision status", policy.DisplayName))
            {
                return;
            }

            if (HandleLockedResourceError(ex, "Check reprovision status", policy.DisplayName))
            {
                return;
            }

            TimedMessage($"[red]Failed to retrieve reprovision status: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (result is null)
        {
            TimedMessage("[yellow]No reprovision status is available yet.[/]");
            return;
        }

        var statusMarkup = result.Status?.ToLowerInvariant() switch
        {
            "succeeded" => "[green]Succeeded[/]",
            "pending" => "[yellow]Pending[/]",
            "failed" => "[red]Failed[/]",
            _ => Markup.Escape(result.Status ?? "-")
        };

        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Reprovision status");
        AnsiConsole.MarkupLine($"[bold]Status[/] {statusMarkup}");
        AnsiConsole.MarkupLine($"[bold]Started[/] {Markup.Escape(result.StartDateTime?.ToLocalTime().ToString("g") ?? "-")}");
        AnsiConsole.MarkupLine($"[bold]Finished[/] {Markup.Escape(result.FinishDateTime?.ToLocalTime().ToString("g") ?? "-")}");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[grey]Press any key to return...[/]");
        ReadNavigationKey(intercept: true);
    }

    /// <summary>
    /// Admin-level rollup of user experience sync (UES) storage across every shared-by-Entra-group
    /// provisioning policy — total/used/remaining GB and profile count per policy, plus an org-wide
    /// aggregate bar. Enter on a row drills into that policy's per-user profile list.
    /// </summary>
    private async Task ShowUserExperienceSyncOverviewAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var policies = await LoadProvisioningPoliciesAsync();
        var uesPolicies = policies.Where(IsSharedByEntraGroupPolicy).ToArray();

        if (uesPolicies.Length == 0)
        {
            TimedMessage("[yellow]No shared-by-Entra-group provisioning policies were found.[/]");
            return;
        }

        var policiesById = uesPolicies.ToDictionary(policy => policy.Id, StringComparer.OrdinalIgnoreCase);

        async Task<IReadOnlyList<GraphTableRow>> LoadOverviewRowsAsync()
        {
            // Each policy's own context -> usage/profiles calls are inherently sequential (usage
            // and profiles both depend on the context lookup that comes first), but different
            // policies are entirely independent of each other, so run policies concurrently
            // (bounded) rather than one full policy-chain at a time.
            var rows = await ConcurrencyHelper.MapWithConcurrencyAsync(uesPolicies, maxConcurrency: 5, async policy =>
            {
                string enabledText;
                string totalText = "-";
                string usedText = "-";
                string remainingText = "-";
                string profileCountText = "-";

                try
                {
                    var context = await _session.Graph.GetUserSettingsPersistenceContextAsync(policy.Id);
                    if (context is null)
                    {
                        enabledText = "Not configured";
                    }
                    else if (!context.Enabled)
                    {
                        enabledText = "Disabled";
                    }
                    else
                    {
                        enabledText = "Enabled";
                        var usage = await _session.Graph.GetUserSettingsPersistenceUsageAsync(context);
                        if (usage is not null)
                        {
                            totalText = (usage.TotalAllocatedStorageInGB ?? 0).ToString("0.#");
                            usedText = (usage.UsedStorageInGB ?? 0).ToString("0.#");
                            remainingText = (usage.RemainingAvailableStorageInGB ?? Math.Max(0, (usage.TotalAllocatedStorageInGB ?? 0) - (usage.UsedStorageInGB ?? 0))).ToString("0.#");
                        }

                        var profiles = await _session.Graph.GetUserSettingsPersistenceProfilesAsync(context);
                        profileCountText = profiles.Count.ToString();
                    }
                }
                catch (Exception ex)
                {
                    enabledText = $"Error: {ex.Message}";
                }

                var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["policyId"] = policy.Id,
                    ["policyName"] = policy.DisplayName,
                    ["status"] = enabledText,
                    ["totalAllocatedStorageInGB"] = totalText,
                    ["usedStorageInGB"] = usedText,
                    ["remainingAvailableStorageInGB"] = remainingText,
                    ["profileCount"] = profileCountText
                };

                return new GraphTableRow(policy.DisplayName, enabledText, fields);
            });

            return rows;
        }

        await ShowGraphRowsAsync(
            "User experience sync overview",
            LoadOverviewRowsAsync,
            GetUserSettingsPersistenceOverviewHeader,
            FormatUserSettingsPersistenceOverviewRow,
            enterAction: async row =>
            {
                var policyId = GetOptionalField(row, "policyId");
                if (policyId is not null && policiesById.TryGetValue(policyId, out var policy))
                {
                    await ShowUserExperienceSyncAsync(policy);
                }
            },
            summaryRenderer: RenderUserSettingsPersistenceOverviewSummary);
    }

    private static void RenderUserSettingsPersistenceOverviewSummary(IReadOnlyList<GraphTableRow> rows)
    {
        double total = 0;
        double used = 0;
        var enabledCount = 0;

        foreach (var row in rows)
        {
            if (!string.Equals(GetField(row, "status"), "Enabled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            enabledCount++;
            if (double.TryParse(GetField(row, "totalAllocatedStorageInGB"), out var rowTotal))
            {
                total += rowTotal;
            }

            if (double.TryParse(GetField(row, "usedStorageInGB"), out var rowUsed))
            {
                used += rowUsed;
            }
        }

        AnsiConsole.MarkupLine($"[bold]Policies with UES enabled[/] {enabledCount} of {rows.Count}");

        if (enabledCount == 0)
        {
            return;
        }

        var remaining = Math.Max(0, total - used);
        AnsiConsole.MarkupLine($"[bold]Total allocated[/] {total:0.#} GB    [bold]Used[/] {used:0.#} GB    [bold]Remaining[/] {remaining:0.#} GB");

        const int barWidth = 40;
        var usedSegments = total > 0 ? (int)Math.Round(barWidth * Math.Clamp(used / total, 0, 1)) : 0;
        var freeSegments = Math.Max(0, barWidth - usedSegments);
        var barColor = total > 0 && used / total >= 0.9 ? "red" : total > 0 && used / total >= 0.75 ? "yellow" : "green";
        AnsiConsole.MarkupLine($"[{barColor}]{new string('█', usedSegments)}[/][grey]{new string('░', freeSegments)}[/]");
    }

    private static string GetUserSettingsPersistenceOverviewHeader()
    {
        var widths = GetUserSettingsPersistenceOverviewWidths();
        return Row("Policy", widths.Policy, "Status", widths.Status, "Total (GB)", widths.Total, "Used (GB)", widths.Used, "Remaining (GB)", widths.Remaining, "Profiles", widths.Profiles);
    }

    private static string FormatUserSettingsPersistenceOverviewRow(GraphTableRow row)
    {
        var widths = GetUserSettingsPersistenceOverviewWidths();
        return Row(
            GetField(row, "policyName"), widths.Policy,
            GetField(row, "status"), widths.Status,
            GetField(row, "totalAllocatedStorageInGB"), widths.Total,
            GetField(row, "usedStorageInGB"), widths.Used,
            GetField(row, "remainingAvailableStorageInGB"), widths.Remaining,
            GetField(row, "profileCount"), widths.Profiles);
    }

    private static (int Policy, int Status, int Total, int Used, int Remaining, int Profiles) GetUserSettingsPersistenceOverviewWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int status = 14;
        const int total = 11;
        const int used = 10;
        const int remaining = 15;
        const int profiles = 9;
        var gaps = 5;
        var policy = Math.Max(20, available - status - total - used - remaining - profiles - gaps);
        return (policy, status, total, used, remaining, profiles);
    }

    /// <summary>
    /// "Manage group members" — lists members of the Entra group(s) assigned to this provisioning
    /// policy, with actions to add or remove members directly from the CLI (mirrors what an admin
    /// would otherwise do in the Entra/Azure AD portal's group membership blade).
    /// </summary>
    private async Task ShowProvisioningPolicyGroupMembersAsync(ProvisioningPolicySummary policy)
    {
        if (policy.AssignedGroupIds.Count == 0)
        {
            TimedMessage("[yellow]This policy has no assigned group.[/]");
            return;
        }

        string groupId;
        string groupName;

        if (policy.AssignedGroupIds.Count == 1)
        {
            groupId = policy.AssignedGroupIds[0];
            groupName = policy.AssignedGroupNames.Count > 0 ? policy.AssignedGroupNames[0] : groupId;
        }
        else
        {
            var groupChoices = policy.AssignedGroupNames.Append("Back").ToArray();
            var choice = PromptChoice(
                () => RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Manage group members"),
                "Select a group to manage",
                groupChoices,
                "Back");

            if (choice == "Back")
            {
                return;
            }

            var index = policy.AssignedGroupNames.ToList().IndexOf(choice);
            groupId = policy.AssignedGroupIds[Math.Max(0, index)];
            groupName = choice;
        }

        var selectedIndex = 0;

        while (true)
        {
            List<GroupMemberSummary> members;
            try
            {
                members = (await AnsiConsole.Status()
                    .Spinner(Spinner.Known.Dots)
                    .StartAsync($"Loading members of {groupName}...", async _ => await _session.Graph.GetGroupMembersAsync(groupId)))
                    .OrderBy(member => member.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                if (await HandlePermissionErrorAsync(ex, "Manage group members", groupName, "GroupMember.Read.All or Group.ReadWrite.All") ||
                    HandleLockedResourceError(ex, "Manage group members", groupName))
                {
                    return;
                }

                TimedMessage($"[red]Failed to load group members: {Markup.Escape(ex.Message)}[/]");
                return;
            }

            if (selectedIndex >= members.Count)
            {
                selectedIndex = Math.Max(0, members.Count - 1);
            }

            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "Manage group members");
            AnsiConsole.MarkupLine($"[#58a6ff]Group members[/] [grey]{Markup.Escape(groupName)}[/]  [grey]({members.Count})[/]");
            AnsiConsole.WriteLine();

            var table = new Table().Border(TableBorder.Rounded).AddColumn(" ").AddColumn("Name").AddColumn("UPN");
            if (members.Count == 0)
            {
                table.AddRow(" ", "[grey]No members found.[/]", "-");
            }
            else
            {
                for (var index = 0; index < members.Count; index++)
                {
                    var member = members[index];
                    var selected = index == selectedIndex;
                    table.AddRow(
                        selected ? "[black on #58a6ff]>[/]" : " ",
                        selected ? Selected(Markup.Escape(member.DisplayName ?? "-")) : Markup.Escape(member.DisplayName ?? "-"),
                        selected ? Selected(Markup.Escape(member.UserPrincipalName ?? "-")) : Markup.Escape(member.UserPrincipalName ?? "-"));
                }
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine("[grey]Up/Down select | Enter view/remove | A add member | R refresh | Esc/B/Q back[/]");

            var key = ReadNavigationKey(intercept: true);
            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedIndex = Math.Max(0, selectedIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedIndex = Math.Min(Math.Max(0, members.Count - 1), selectedIndex + 1);
                    break;
                case ConsoleKey.Home:
                    selectedIndex = 0;
                    break;
                case ConsoleKey.End:
                    selectedIndex = Math.Max(0, members.Count - 1);
                    break;
                case ConsoleKey.Enter:
                    if (members.Count > 0)
                    {
                        await ShowGroupMemberDetailAsync(groupId, groupName, members[selectedIndex]);
                    }
                    break;
                case ConsoleKey.Escape:
                case ConsoleKey.LeftArrow:
                    return;
                default:
                    if (key.KeyChar is 'a' or 'A')
                    {
                        await AddGroupMemberWizardAsync(groupId, groupName);
                    }
                    else if (key.KeyChar is 'r' or 'R')
                    {
                        // Loop reloads members at the top — nothing else to do here.
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
    /// Detail/actions screen for a single group member — Enter never removes directly; the user
    /// must arrow down to "Remove" and press Enter again, avoiding an accidental destructive action.
    /// </summary>
    private async Task ShowGroupMemberDetailAsync(string groupId, string groupName, GroupMemberSummary member)
    {
        var actions = new[] { "Remove from group", "Back" };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Provisioning", "Policies", groupName, "Member");
            var details = new Panel(new Rows(
                    new Markup(PropertyInline("Name", member.DisplayName ?? "-")),
                    new Markup(PropertyInline("UPN", member.UserPrincipalName ?? "-")),
                    new Markup(PropertyBlock("User ID", member.Id))))
                .Header("Member details")
                .Border(BoxBorder.Rounded);

            var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
            var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
                .Header("Actions")
                .Border(BoxBorder.Rounded);

            if (Console.WindowWidth >= 120)
            {
                var grid = new Grid();
                grid.AddColumn();
                grid.AddColumn();
                grid.AddRow(details, actionPanel);
                AnsiConsole.Write(grid);
            }
            else
            {
                AnsiConsole.Write(details);
                AnsiConsole.Write(actionPanel);
            }

            AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back[/]");
            var key = ReadNavigationKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedActionIndex = Math.Max(0, selectedActionIndex - 1);
                    break;
                case ConsoleKey.DownArrow:
                    selectedActionIndex = Math.Min(actions.Length - 1, selectedActionIndex + 1);
                    break;
                case ConsoleKey.Enter:
                    if (actions[selectedActionIndex] == "Back")
                    {
                        return;
                    }

                    await ConfirmAndRunAsync(
                        "Remove from group",
                        $"{member.Name} from {groupName}",
                        async () => await _session.Graph.RemoveGroupMemberAsync(groupId, member.Id),
                        resourceType: "Group member",
                        resourceName: member.Name,
                        requiredPermission: "GroupMember.ReadWrite.All or Group.ReadWrite.All");
                    return;
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

    /// <summary>
    /// Search-then-pick wizard for adding a user to the policy's assigned group. Prompts for a
    /// search term (name/UPN/email prefix), shows matches, then adds the selected user via Graph.
    /// </summary>
    private async Task AddGroupMemberWizardAsync(string groupId, string groupName)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", groupName, "Add member");
        AnsiConsole.MarkupLine($"[#58a6ff]Add member[/] [grey]{Markup.Escape(groupName)}[/]");
        AnsiConsole.WriteLine();

        var query = AnsiConsole.Ask<string>("Search by name, UPN, or email:");
        if (string.IsNullOrWhiteSpace(query))
        {
            return;
        }

        List<GroupMemberSummary> matches;
        try
        {
            matches = (await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Searching directory...", async _ => await _session.Graph.SearchUsersAsync(query)))
                .ToList();
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Add member", groupName, "User.Read.All or Directory.Read.All") ||
                HandleLockedResourceError(ex, "Add member", groupName))
            {
                return;
            }

            TimedMessage($"[red]Failed to search directory: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (matches.Count == 0)
        {
            TimedMessage("[yellow]No matching users were found.[/]");
            return;
        }

        var choiceLabels = matches
            .Select(match => $"{match.DisplayName ?? "-"}  <{match.UserPrincipalName ?? "-"}>")
            .Append("Cancel")
            .ToArray();

        var selected = PromptChoice(
            () => RenderBreadcrumb("Provisioning", "Policies", groupName, "Add member"),
            "Select a user to add",
            choiceLabels,
            "Cancel");

        if (selected == "Cancel")
        {
            return;
        }

        var selectedIndex = Array.IndexOf(choiceLabels, selected);
        var user = matches[selectedIndex];

        await ConfirmAndRunAsync(
            "Add group member",
            $"{user.Name} to {groupName}",
            async () => await _session.Graph.AddGroupMemberAsync(groupId, user.Id),
            resourceType: "Group member",
            resourceName: user.Name,
            requiredPermission: "GroupMember.ReadWrite.All or Group.ReadWrite.All");
    }

    /// <summary>
    /// "User experience sync" — surfaces the user settings persistence storage usage bar and
    /// per-user profile list for a shared-by-Entra-group provisioning policy, matching the view
    /// the Intune portal shows under a shared policy's assignment details.
    /// </summary>
    private async Task ShowUserExperienceSyncAsync(ProvisioningPolicySummary policy)
    {
        AnsiConsole.Clear();
        RenderBreadcrumb("Provisioning", "Policies", policy.DisplayName, "User experience sync");

        ProvisioningPolicyUserSettingsPersistenceContext? context;
        try
        {
            context = await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Resolving user experience sync configuration...", async _ =>
                    await _session.Graph.GetUserSettingsPersistenceContextAsync(policy.Id));
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "User experience sync", policy.DisplayName))
            {
                return;
            }

            if (HandleLockedResourceError(ex, "User experience sync", policy.DisplayName))
            {
                return;
            }

            TimedMessage($"[red]Failed to resolve user experience sync configuration: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        if (context is null)
        {
            TimedMessage("[yellow]No assignment with user experience sync (user settings persistence) was found for this policy.[/]");
            return;
        }

        if (!context.Enabled)
        {
            TimedMessage("[yellow]User experience sync is not enabled on this policy's assignment.[/]");
            return;
        }

        CloudPcUserSettingsPersistenceUsageResult? usage = null;

        async Task<IReadOnlyList<GraphTableRow>> LoadProfilesAsync()
        {
            // Refresh usage alongside the profile list (both on initial load and on manual "R"
            // refresh) so the usage bar reflects reality after a profile is deleted.
            try
            {
                usage = await _session.Graph.GetUserSettingsPersistenceUsageAsync(context);
            }
            catch
            {
                usage = null;
            }

            return await _session.Graph.GetUserSettingsPersistenceProfilesAsync(context);
        }

        await ShowGraphRowsAsync(
            "User profiles",
            LoadProfilesAsync,
            GetUserSettingsPersistenceProfilesHeader,
            FormatUserSettingsPersistenceProfileRow,
            enterAction: row => ShowUserSettingsPersistenceProfileDetailAsync(context, row),
            summaryRenderer: _ =>
            {
                if (usage is not null)
                {
                    RenderUserSettingsPersistenceUsageBar(usage);
                }
                else
                {
                    AnsiConsole.MarkupLine("[yellow]Storage usage is not available.[/]");
                }
            });
    }

    /// <summary>
    /// Detail/actions screen for a single UES profile row (opened via Enter on the profile list).
    /// Shows the profile's fields and offers "Delete" as an explicit, arrow-key-selected action —
    /// Enter on the row itself never deletes directly, avoiding an accidental destructive action.
    /// </summary>
    private async Task ShowUserSettingsPersistenceProfileDetailAsync(ProvisioningPolicyUserSettingsPersistenceContext context, GraphTableRow row)
    {
        var actions = new[] { "Delete", "Back" };
        var selectedActionIndex = 0;

        while (true)
        {
            AnsiConsole.Clear();
            RenderUserSettingsPersistenceProfileDetailLayout(row, actions, selectedActionIndex);
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

                    await DeleteUserSettingsPersistenceProfileAsync(context, row);
                    return;
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

    private static void RenderUserSettingsPersistenceProfileDetailLayout(GraphTableRow row, IReadOnlyList<string> actions, int selectedActionIndex)
    {
        RenderBreadcrumb("Provisioning", "Policies", "User experience sync", "Profile");
        var details = new Panel(new Rows(
                new Markup(PropertyInline("User", GetField(row, "userPrincipalName"))),
                new Markup(PropertyInline("Size (GB)", GetField(row, "profileSizeInGB"))),
                new Markup(PropertyInline("Status", GetField(row, "status"))),
                new Markup(PropertyInline("Last attached", GetField(row, "lastProfileAttachedDateTime"))),
                new Markup(PropertyBlock("Profile ID", GetField(row, "profileId")))))
            .Header("Profile details")
            .Border(BoxBorder.Rounded);

        var actionLines = actions.Select((action, index) => FormatActionLine(action, index == selectedActionIndex));
        var actionPanel = new Panel(new Markup(string.Join(Environment.NewLine, actionLines)))
            .Header("Actions")
            .Border(BoxBorder.Rounded);

        if (Console.WindowWidth >= 120)
        {
            var grid = new Grid();
            grid.AddColumn();
            grid.AddColumn();
            grid.AddRow(details, actionPanel);
            AnsiConsole.Write(grid);
        }
        else
        {
            AnsiConsole.Write(details);
            AnsiConsole.Write(actionPanel);
        }

        AnsiConsole.MarkupLine("[grey]Up/Down choose action | Enter run | Esc/B/Q back[/]");
    }

    /// <summary>
    /// Deletes ("cleans up") a UES disk/profile for one user, after the "Delete" action is chosen
    /// on the profile detail screen. Mirrors the Intune portal's per-profile delete, which puts the
    /// profile into a "deleting" status the next time the list is refreshed (press R).
    /// </summary>
    private async Task DeleteUserSettingsPersistenceProfileAsync(ProvisioningPolicyUserSettingsPersistenceContext context, GraphTableRow row)
    {
        var profileId = GetOptionalField(row, "profileId", "ProfileId");
        var upn = GetOptionalField(row, "userPrincipalName", "UserPrincipalName") ?? "this user";

        if (string.IsNullOrWhiteSpace(profileId))
        {
            TimedMessage("[yellow]This profile has no profileId — it cannot be deleted.[/]");
            return;
        }

        var deleted = false;
        await ConfirmAndRunAsync(
            "Delete user experience sync profile",
            upn,
            async () =>
            {
                await _session.Graph.BatchCleanupUserSettingsPersistenceProfileAsync(context, profileId);
                deleted = true;
            },
            resourceType: "UES profile",
            resourceName: upn);

        if (deleted)
        {
            TimedMessage("[grey]Deletion queued. Press R on the profile list to refresh status.[/]");
        }
    }

    private static void RenderUserSettingsPersistenceUsageBar(CloudPcUserSettingsPersistenceUsageResult usage)
    {
        var total = usage.TotalAllocatedStorageInGB ?? 0;
        var used = usage.UsedStorageInGB ?? 0;
        var remaining = usage.RemainingAvailableStorageInGB ?? Math.Max(0, total - used);

        AnsiConsole.MarkupLine($"[bold]Total allocated[/] {total:0.#} GB    [bold]Used[/] {used:0.#} GB    [bold]Remaining[/] {remaining:0.#} GB");

        const int barWidth = 40;
        var usedSegments = total > 0 ? (int)Math.Round(barWidth * Math.Clamp(used / total, 0, 1)) : 0;
        var freeSegments = Math.Max(0, barWidth - usedSegments);
        var barColor = total > 0 && used / total >= 0.9 ? "red" : total > 0 && used / total >= 0.75 ? "yellow" : "green";
        AnsiConsole.MarkupLine($"[{barColor}]{new string('█', usedSegments)}[/][grey]{new string('░', freeSegments)}[/]");
    }

    private static string GetUserSettingsPersistenceProfilesHeader()
    {
        var widths = GetUserSettingsPersistenceProfileWidths();
        return Row("UPN", widths.Upn, "Size (GB)", widths.Size, "Status", widths.Status, "Last attached", widths.LastAttached);
    }

    private static string FormatUserSettingsPersistenceProfileRow(GraphTableRow row)
    {
        var widths = GetUserSettingsPersistenceProfileWidths();
        return Row(
            GetField(row, "userPrincipalName"), widths.Upn,
            GetField(row, "profileSizeInGB"), widths.Size,
            GetField(row, "status"), widths.Status,
            GetField(row, "lastProfileAttachedDateTime"), widths.LastAttached);
    }

    private static (int Upn, int Size, int Status, int LastAttached) GetUserSettingsPersistenceProfileWidths()
    {
        var available = Math.Max(76, Console.WindowWidth - 4);
        const int size = 10;
        const int status = 14;
        const int lastAttached = 20;
        var gaps = 3;
        var upn = Math.Max(20, available - size - status - lastAttached - gaps);
        return (upn, size, status, lastAttached);
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(character => invalid.Contains(character) ? '-' : character));
    }

    private static string FormatBool(bool? value)
    {
        return value is null ? "-" : value.Value ? "Yes" : "No";
    }

    private static IReadOnlyList<ProvisioningPolicySummary> FilterProvisioningPolicies(IReadOnlyList<ProvisioningPolicySummary> policies, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return policies;
        }

        return policies
            .Where(policy =>
                Contains(policy.DisplayName, filter) ||
                Contains(policy.Description, filter) ||
                Contains(policy.ProvisioningType, filter) ||
                Contains(policy.ImageDisplayName, filter) ||
                Contains(policy.ImageType, filter) ||
                Contains(policy.DomainJoinTypes, filter) ||
                Contains(policy.CloudPcNamingTemplate, filter) ||
                Contains(policy.CloudPcGroupDisplayName, filter) ||
                policy.AssignedGroupNames.Any(groupName => Contains(groupName, filter)))
            .ToArray();
    }

    private static IReadOnlyList<ProvisioningPolicySummary> SortProvisioningPolicies(IReadOnlyList<ProvisioningPolicySummary> policies, ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Type => policies.OrderBy(policy => policy.ProvisioningType, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            ProvisioningPolicySortMode.Image => policies.OrderBy(policy => policy.ImageDisplayName, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            ProvisioningPolicySortMode.Join => policies.OrderBy(policy => policy.DomainJoinTypes, StringComparer.OrdinalIgnoreCase).ThenBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray(),
            _ => policies.OrderBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray()
        };
    }

    private static ProvisioningPolicySortMode NextProvisioningPolicySortMode(ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Name => ProvisioningPolicySortMode.Type,
            ProvisioningPolicySortMode.Type => ProvisioningPolicySortMode.Image,
            ProvisioningPolicySortMode.Image => ProvisioningPolicySortMode.Join,
            _ => ProvisioningPolicySortMode.Name
        };
    }

    private static string FormatProvisioningPolicySortMode(ProvisioningPolicySortMode sortMode)
    {
        return sortMode switch
        {
            ProvisioningPolicySortMode.Type => "type",
            ProvisioningPolicySortMode.Image => "image",
            ProvisioningPolicySortMode.Join => "join",
            _ => "name"
        };
    }

    private enum ProvisioningPolicySortMode
    {
        Name,
        Type,
        Image,
        Join
    }
}
