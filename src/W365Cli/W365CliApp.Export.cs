using Spectre.Console;
using System.Text;

namespace W365Cli;

internal sealed partial class W365CliApp
{
    /// <summary>
    /// Generates a single Markdown document summarizing the current tenant state -- Cloud PC
    /// inventory, provisioning policies, and (if the signed-in account has the extra directory
    /// licensing permissions) Windows 365 license consumption -- for sharing or archiving outside
    /// the CLI. Reuses the exact same data sources and calculations already used by Browse Cloud
    /// PCs, Provisioning Policies, and Licensing, so figures here always match what those live
    /// screens show. Reads-only; makes no changes to the tenant.
    /// </summary>
    private async Task ExportMarkdownSnapshotAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        AnsiConsole.Clear();
        RenderBreadcrumb("Export");
        var defaultPath = Path.Combine(Environment.CurrentDirectory, $"W365-Snapshot-{DateTime.Now:yyyy-MM-dd-HHmm}.md");
        var path = PromptTextCancelable($"Export path [[{Markup.Escape(defaultPath)}]]:", allowEmpty: true);
        if (path is null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            path = defaultPath;
        }

        IReadOnlyList<CloudPcSummary> cloudPcs = [];
        IReadOnlyList<ProvisioningPolicySummary> policies = [];
        IReadOnlyList<LicenseOverviewItem> licenseItems = [];
        string? licensingError = null;

        try
        {
            await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Gathering tenant data...", async _ =>
                {
                    cloudPcs = await _session.Graph.GetCloudPcsAsync();
                    policies = await _session.Graph.GetProvisioningPoliciesAsync();

                    // Licensing needs extra directory permissions (subscribedSkus) that not every
                    // signed-in account has -- unlike Cloud PCs/Policies above, a failure here
                    // shouldn't block the rest of the export, just note it and move on, matching
                    // ShowLicensingAsync's own error handling for the interactive Licensing screen.
                    try
                    {
                        var skus = await _session.Graph.GetSubscribedSkusAsync();
                        licenseItems = BuildLicenseOverview(skus, cloudPcs, policies);
                    }
                    catch (Exception ex)
                    {
                        licensingError = ex.Message;
                    }
                });
        }
        catch (Exception ex)
        {
            if (await HandlePermissionErrorAsync(ex, "Export Markdown Snapshot", "tenant data") ||
                HandleLockedResourceError(ex, "Export Markdown Snapshot", "tenant data"))
            {
                return;
            }

            TimedMessage($"[red]Failed to gather tenant data: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        var markdown = BuildMarkdownSnapshot(cloudPcs, policies, licenseItems, licensingError);

        try
        {
            await File.WriteAllTextAsync(path, markdown);
        }
        catch (Exception ex)
        {
            TimedMessage($"[red]Failed to write export file: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        ShowActionResult("Exported", "Export Markdown Snapshot", path, "[green]Exported.[/]");
    }

    /// <summary>
    /// Exports every currently loaded row of a Graph rows screen (Sign-in status, Launch details,
    /// etc.) to a CSV file -- not just the visible/filtered page. When <paramref name="csvColumns"/>
    /// is supplied, uses that explicit, ordered header/field list (so a report can pick a
    /// user-friendly subset and order, e.g. the Last Sign-In Report); otherwise falls back to the
    /// union of every row's Fields keys, sorted, so any Graph rows screen can be exported even
    /// without a bespoke column list.
    /// </summary>
    private async Task ExportGraphRowsToCsvAsync(
        string title,
        IReadOnlyList<GraphTableRow> rows,
        IReadOnlyList<(string Header, string Field)>? csvColumns)
    {
        if (rows.Count == 0)
        {
            TimedMessage("[yellow]No rows to export.[/]");
            return;
        }

        AnsiConsole.Clear();
        RenderBreadcrumb("Export CSV");
        var safeTitle = string.Concat(title.Select(ch => char.IsLetterOrDigit(ch) ? ch : '-'));
        var defaultPath = Path.Combine(Environment.CurrentDirectory, $"W365-{safeTitle}-{DateTime.Now:yyyy-MM-dd-HHmm}.csv");
        var path = PromptTextCancelable($"Export path [[{Markup.Escape(defaultPath)}]]:", allowEmpty: true);
        if (path is null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            path = defaultPath;
        }

        var columns = csvColumns ?? rows
            .SelectMany(row => row.Fields.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
            .Select(key => (Header: key, Field: key))
            .ToArray();

        var csv = BuildCsv(columns, rows);

        try
        {
            await File.WriteAllTextAsync(path, csv);
        }
        catch (Exception ex)
        {
            TimedMessage($"[red]Failed to write export file: {Markup.Escape(ex.Message)}[/]");
            return;
        }

        ShowActionResult("Exported", "Export CSV", path, $"[green]Exported {rows.Count} row(s).[/]");
    }

    internal static string BuildCsv(IReadOnlyList<(string Header, string Field)> columns, IReadOnlyList<GraphTableRow> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", columns.Select(column => CsvEscape(column.Header))));
        foreach (var row in rows)
        {
            sb.AppendLine(string.Join(",", columns.Select(column =>
                CsvEscape(row.Fields.TryGetValue(column.Field, out var value) ? value : null))));
        }

        return sb.ToString();
    }

    internal static string CsvEscape(string? value)
    {
        value ??= string.Empty;
        return value.Contains(',', StringComparison.Ordinal) ||
               value.Contains('"', StringComparison.Ordinal) ||
               value.Contains('\n', StringComparison.Ordinal) ||
               value.Contains('\r', StringComparison.Ordinal)
            ? $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\""
            : value;
    }

    private string BuildMarkdownSnapshot(
        IReadOnlyList<CloudPcSummary> cloudPcs,
        IReadOnlyList<ProvisioningPolicySummary> policies,
        IReadOnlyList<LicenseOverviewItem> licenseItems,
        string? licensingError)
    {
        var sb = new StringBuilder();
        var tenantLabel = _session.TenantName ?? _session.TenantId ?? "Unknown tenant";

        sb.AppendLine("# Windows 365 Tenant Snapshot");
        sb.AppendLine();
        sb.AppendLine($"- **Generated:** {DateTimeOffset.Now:f}");
        sb.AppendLine($"- **Tenant:** {MarkdownEscape(tenantLabel)}");
        if (!string.IsNullOrWhiteSpace(_session.SignedInUserUpn))
        {
            sb.AppendLine($"- **Signed in as:** {MarkdownEscape(_session.SignedInUserUpn)}");
        }

        sb.AppendLine($"- **Generated by:** W365 CLI v{GetCurrentVersion()}");
        sb.AppendLine();

        AppendCloudPcSection(sb, cloudPcs);
        AppendProvisioningPolicySection(sb, policies);
        AppendLicensingSection(sb, licenseItems, licensingError);

        return sb.ToString();
    }

    internal static void AppendCloudPcSection(StringBuilder sb, IReadOnlyList<CloudPcSummary> cloudPcs)
    {
        sb.AppendLine("## Cloud PCs");
        sb.AppendLine();
        sb.AppendLine($"Total: **{cloudPcs.Count}**");
        sb.AppendLine();

        if (cloudPcs.Count == 0)
        {
            sb.AppendLine("_No Cloud PCs were returned._");
            sb.AppendLine();
            return;
        }

        var statusSummary = cloudPcs
            .GroupBy(pc => pc.Status ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}");
        sb.AppendLine($"**By status:** {string.Join(", ", statusSummary)}");
        sb.AppendLine();

        var typeSummary = cloudPcs
            .GroupBy(pc => pc.ProvisioningType ?? "unknown", StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => $"{group.Key}: {group.Count()}");
        sb.AppendLine($"**By type:** {string.Join(", ", typeSummary)}");
        sb.AppendLine();

        AppendMarkdownTable(
            sb,
            ["Name", "Status", "Type", "User", "Service plan"],
            cloudPcs
                .OrderBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase)
                .Select(pc => new[]
                {
                    MarkdownEscape(pc.Name),
                    MarkdownEscape(pc.Status),
                    MarkdownEscape(pc.ProvisioningType),
                    MarkdownEscape(pc.EffectiveUserPrincipalName),
                    MarkdownEscape(pc.ServicePlanName)
                }));
    }

    internal static void AppendProvisioningPolicySection(StringBuilder sb, IReadOnlyList<ProvisioningPolicySummary> policies)
    {
        sb.AppendLine("## Provisioning Policies");
        sb.AppendLine();
        sb.AppendLine($"Total: **{policies.Count}**");
        sb.AppendLine();

        if (policies.Count == 0)
        {
            sb.AppendLine("_No provisioning policies were returned._");
            sb.AppendLine();
            return;
        }

        AppendMarkdownTable(
            sb,
            ["Name", "Type", "Image", "Domain join", "SSO", "Assigned groups"],
            policies
                .OrderBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase)
                .Select(policy => new[]
                {
                    MarkdownEscape(policy.DisplayName),
                    MarkdownEscape(policy.ProvisioningType),
                    MarkdownEscape(policy.ImageDisplayName),
                    MarkdownEscape(policy.DomainJoinTypes),
                    FormatBool(policy.EnableSingleSignOn),
                    MarkdownEscape(policy.AssignedGroupNames.Count == 0 ? null : string.Join(", ", policy.AssignedGroupNames))
                }));
    }

    internal static void AppendLicensingSection(StringBuilder sb, IReadOnlyList<LicenseOverviewItem> licenseItems, string? licensingError)
    {
        sb.AppendLine("## Licensing");
        sb.AppendLine();

        if (licensingError is not null)
        {
            sb.AppendLine("_Licensing data could not be loaded for this export (requires subscribedSkus / directory licensing read permissions)._");
            sb.AppendLine();
            sb.AppendLine($"> {MarkdownEscape(licensingError)}");
            sb.AppendLine();
            return;
        }

        if (licenseItems.Count == 0)
        {
            sb.AppendLine("_No Windows 365 license SKUs were detected._");
            sb.AppendLine();
            return;
        }

        AppendMarkdownTable(
            sb,
            ["Family", "SKUs", "Purchased", "Assigned", "Cloud PCs", "Dedicated", "Shared", "Units used", "Units left"],
            licenseItems.Select(item => new[]
            {
                MarkdownEscape(item.Family),
                MarkdownEscape(item.SkuPartNumbers),
                item.Purchased.ToString(),
                item.Assigned.ToString(),
                item.CloudPcCount.ToString(),
                item.DedicatedCloudPcCount.ToString(),
                item.SharedCloudPcCount.ToString(),
                item.LicenseUnitsUsed.ToString(),
                item.LicenseUnitsLeft.ToString()
            }));
    }

    internal static string MarkdownEscape(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "-";
        }

        return value.Replace("|", "\\|", StringComparison.Ordinal).Replace("\r", "", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal);
    }

    internal static void AppendMarkdownTable(StringBuilder sb, IReadOnlyList<string> headers, IEnumerable<IReadOnlyList<string>> rows)
    {
        sb.AppendLine($"| {string.Join(" | ", headers)} |");
        sb.AppendLine($"| {string.Join(" | ", headers.Select(_ => "---"))} |");
        foreach (var row in rows)
        {
            sb.AppendLine($"| {string.Join(" | ", row)} |");
        }

        sb.AppendLine();
    }
}
