using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

        private async Task ShowLicensingAsync()
        {
            if (!await EnsureConnectedAsync())
            {
                return;
            }

            IReadOnlyList<LicenseOverviewItem> items;
            try
            {
                items = await LoadLicenseOverviewAsync();
            }
            catch (Exception ex)
            {
                AnsiConsole.Clear();
                RenderBreadcrumb("Licensing");
                AnsiConsole.MarkupLine("[red]Failed to load licensing data.[/]");
                AnsiConsole.MarkupLine("[grey]The Licensing view requires access to subscribedSkUs, usually via Organization.Read.All or equivalent directory licensing permissions.[/]");
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[grey]{Markup.Escape(ex.Message)}[/]");
                WaitForBack();
                return;
            }

            if (items.Count == 0)
            {
                TimedMessage("[yellow]No Windows 365 license SKUs were detected from subscribedSkUs.[/]");
                return;
            }

            var selectedIndex = 0;
            while (true)
            {
                AnsiConsole.Clear();
                RenderLicensingOverview(items, selectedIndex);
                var key = ReadNavigationKey(intercept: true);
                switch (key.Key)
                {
                    case ConsoleKey.UpArrow:
                        selectedIndex = Math.Max(0, selectedIndex - 1);
                        break;
                    case ConsoleKey.DownArrow:
                        selectedIndex = Math.Min(items.Count - 1, selectedIndex + 1);
                        break;
                    case ConsoleKey.PageUp:
                        selectedIndex = Math.Max(0, selectedIndex - 10);
                        break;
                    case ConsoleKey.PageDown:
                        selectedIndex = Math.Min(items.Count - 1, selectedIndex + 10);
                        break;
                    case ConsoleKey.Home:
                        selectedIndex = 0;
                        break;
                    case ConsoleKey.End:
                        selectedIndex = items.Count - 1;
                        break;
                    case ConsoleKey.Enter:
                        await ShowLicenseDetailsAsync(items[selectedIndex]);
                        break;
                    case ConsoleKey.R:
                        items = await LoadLicenseOverviewAsync();
                        selectedIndex = Math.Min(selectedIndex, Math.Max(0, items.Count - 1));
                        if (items.Count == 0)
                        {
                            TimedMessage("[yellow]No Windows 365 license SKUs were detected from subscribedSkUs.[/]");
                            return;
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

        private async Task<IReadOnlyList<LicenseOverviewItem>> LoadLicenseOverviewAsync()
        {
            return await AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .StartAsync("Loading license data...", async _ =>
                {
                    var skus = await _session.Graph.GetSubscribedSkusAsync();
                    var cloudPcs = await _session.Graph.GetCloudPcsAsync();
                    var policies = await _session.Graph.GetProvisioningPoliciesAsync();
                    return BuildLicenseOverview(skus, cloudPcs, policies);
                });
        }

        private static IReadOnlyList<LicenseOverviewItem> BuildLicenseOverview(
            IReadOnlyList<SubscribedSku> skus,
            IReadOnlyList<CloudPcSummary> cloudPcs,
            IReadOnlyList<ProvisioningPolicySummary> policies)
        {
            var windows365Skus = skus
                .Select(sku => new { Sku = sku, Info = GetWindows365LicenseInfo(sku) })
                .Where(item => item.Info is not null)
                .ToArray();
            var output = new List<LicenseOverviewItem>();
            foreach (var group in windows365Skus
                .GroupBy(item => $"{item.Info!.Family}|{item.Info.PlanKey}", StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.First().Info!.Family, StringComparer.OrdinalIgnoreCase)
                .ThenBy(group => group.First().Info!.PlanKey, StringComparer.OrdinalIgnoreCase))
            {
                var info = group.First().Info!;
                var family = info.Family;
                var purchased = group.Sum(item => item.Sku.PrepaidUnits?.Enabled ?? 0);
                var assigned = group.Sum(item => item.Sku.ConsumedUnits ?? 0);
                var isFlex = family.Equals("Flex", StringComparison.OrdinalIgnoreCase);
                var matchingCloudPcs = GetCloudPcsForLicense(cloudPcs, info);
                var dedicated = isFlex ? matchingCloudPcs.Count(pc => GetFlexCloudPcMode(pc) == "Dedicated") : 0;
                var shared = isFlex ? matchingCloudPcs.Count - dedicated : 0;
                var dedicatedUnitsUsed = isFlex ? (int)Math.Ceiling(dedicated / 3d) : 0;
                var sharedUnitsUsed = isFlex ? shared : 0;
                var licenseUnitsUsed = isFlex ? sharedUnitsUsed + dedicatedUnitsUsed : matchingCloudPcs.Count;
                var licenseUnitsLeft = Math.Max(0, purchased - licenseUnitsUsed);
                var provisionable = isFlex ? purchased * 3 : purchased;
                var activeLimit = isFlex ? purchased : purchased;
                var available = isFlex
                    ? licenseUnitsLeft * 3 + Math.Max(0, dedicatedUnitsUsed * 3 - dedicated)
                    : Math.Max(0, provisionable - matchingCloudPcs.Count);
                var flexPolicies = policies
                    .Where(policy => isFlex && IsFlexPolicy(policy))
                    .ToArray();

                output.Add(new LicenseOverviewItem(
                    info.DisplayName,
                    string.Join(", ", group.Select(item => item.Sku.SkuPartNumber).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase)),
                    purchased,
                    assigned,
                    matchingCloudPcs.Count,
                    dedicated,
                    shared,
                    provisionable,
                    available,
                    activeLimit,
                    dedicatedUnitsUsed,
                    sharedUnitsUsed,
                    licenseUnitsUsed,
                    licenseUnitsLeft,
                    matchingCloudPcs,
                    flexPolicies));
            }

            return output;
        }

        private static Windows365LicenseInfo? GetWindows365LicenseInfo(SubscribedSku sku)
        {
            var text = $"{sku.SkuPartNumber} {string.Join(' ', sku.ServicePlans?.Select(plan => plan.ServicePlanName) ?? [])}";
            if (!text.Contains("W365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("WINDOWS_365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("WINDOWS365", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CPC_", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CLOUD_PC", StringComparison.OrdinalIgnoreCase) &&
                !text.Contains("CLOUDPC", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (text.Contains("DISASTER", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("ADD_ON", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("ADDON", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var family = IsReserveText(text)
                ? "Reserve"
                : text.Contains("FLEX", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("FRONTLINE", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("CPC_F_", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("CPC_S_", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("WINDOWS_365_S_", StringComparison.OrdinalIgnoreCase)
                    ? "Flex"
                    : text.Contains("BUSINESS", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("CPC_B_", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("WINDOWS_365_B_", StringComparison.OrdinalIgnoreCase)
                        ? "Business"
                        : "Enterprise";
            var planKey = GetPlanKey(text) ?? "unknown";
            var planLabel = planKey == "unknown" ? "unknown size" : FormatPlanKey(planKey);
            return new Windows365LicenseInfo(family, planKey, $"{family} {planLabel}");
        }

        private static IReadOnlyList<CloudPcSummary> GetCloudPcsForLicense(IReadOnlyList<CloudPcSummary> cloudPcs, Windows365LicenseInfo license)
        {
            return cloudPcs
                .Where(pc =>
                    (license.PlanKey.Equals("unknown", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(GetPlanKey(pc.ServicePlanName ?? string.Empty), license.PlanKey, StringComparison.OrdinalIgnoreCase)) &&
                    IsCloudPcInLicenseFamily(pc, license.Family))
                .ToArray();
        }

        private static bool IsCloudPcInLicenseFamily(CloudPcSummary pc, string family)
        {
            if (family.Equals("Flex", StringComparison.OrdinalIgnoreCase))
            {
                return Contains(pc.ServicePlanName, "Frontline") ||
                    Contains(pc.ServicePlanName, "Flex") ||
                    Contains(pc.ProvisioningPolicyName, "Flex");
            }

            if (family.Equals("Reserve", StringComparison.OrdinalIgnoreCase))
            {
                return IsReserveText($"{pc.ServicePlanName} {pc.ProvisioningType} {pc.ProvisioningPolicyName}");
            }

            return Contains(pc.ServicePlanName, family);
        }

        private static bool IsReserveText(string? value)
        {
            return Contains(value, "Reserve") ||
                Contains(value, "CPC_R_") ||
                Contains(value, "WINDOWS_365_R_");
        }

        private static string? GetPlanKey(string value)
        {
            var displayMatch = System.Text.RegularExpressions.Regex.Match(
                value,
                @"(?<cpu>\d+)\s*vCPU[^\d]+(?<ram>\d+)\s*GB[^\d]+(?<storage>\d+)\s*GB",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (displayMatch.Success)
            {
                return $"{displayMatch.Groups["cpu"].Value}/{displayMatch.Groups["ram"].Value}/{displayMatch.Groups["storage"].Value}";
            }

            var skuMatch = System.Text.RegularExpressions.Regex.Match(
                value,
                @"(?<cpu>\d+)\s*C[^\d]+(?<ram>\d+)\s*GB[^\d]+(?<storage>\d+)\s*GB",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return skuMatch.Success
                ? $"{skuMatch.Groups["cpu"].Value}/{skuMatch.Groups["ram"].Value}/{skuMatch.Groups["storage"].Value}"
                : null;
        }

        private static string FormatPlanKey(string planKey)
        {
            var parts = planKey.Split('/');
            return parts.Length == 3 ? $"{parts[0]}vCPU/{parts[1]}GB/{parts[2]}GB" : planKey;
        }

        private static bool IsDedicatedCloudPc(CloudPcSummary cloudPc)
        {
            return cloudPc.ProvisioningType?.Contains("dedicated", StringComparison.OrdinalIgnoreCase) == true ||
                cloudPc.ProvisioningPolicyName?.Contains("dedicated", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static bool IsSharedCloudPc(CloudPcSummary cloudPc)
        {
            return cloudPc.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true ||
                cloudPc.ProvisioningPolicyName?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static bool IsFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true ||
                policy.DisplayName.Contains("Flex", StringComparison.OrdinalIgnoreCase) ||
                policy.Description?.Contains("Flex", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static void RenderLicensingOverview(IReadOnlyList<LicenseOverviewItem> items, int selectedIndex)
        {
            RenderBreadcrumb("Licensing");
            AnsiConsole.MarkupLine("[#58a6ff]Windows 365 licensing[/]");
            AnsiConsole.MarkupLine("[grey]Capacity estimates use Microsoft Graph subscribedSkUs plus current Cloud PC inventory.[/]");
            AnsiConsole.WriteLine();
            var table = new Table()
                .Border(TableBorder.Rounded)
                .AddColumn(" ")
                .AddColumn("Family")
                .AddColumn("Purchased")
                .AddColumn("Assigned")
                .AddColumn("Cloud PCs")
                .AddColumn("Dedicated")
                .AddColumn("Shared")
                .AddColumn("Units used")
                .AddColumn("Units left")
                .AddColumn("Can run now");

            for (var index = 0; index < items.Count; index++)
            {
                var item = items[index];
                table.AddRow(
                    index == selectedIndex ? "[black on #58a6ff]>[/]" : " ",
                    Markup.Escape(item.Family),
                    item.Purchased.ToString(),
                    item.Assigned.ToString(),
                    item.CloudPcCount.ToString(),
                    item.DedicatedCloudPcCount.ToString(),
                    item.SharedCloudPcCount.ToString(),
                    item.LicenseUnitsUsed.ToString(),
                    item.LicenseUnitsLeft.ToString(),
                    item.ActiveSessionLimit.ToString());
            }

            AnsiConsole.Write(table);
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[grey]Enter details | R refresh | Esc/B/Q back[/]");
        }

        private async Task ShowLicenseDetailsAsync(LicenseOverviewItem item)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Licensing", item.Family);
            var detailRows = new List<Markup>
            {
                new(PropertyInline("Family", item.Family)),
                new(PropertyBlock("SKUs", item.SkuPartNumbers)),
                new(PropertyInline("Purchased licenses", item.Purchased.ToString())),
                new(PropertyInline("Assigned licenses", item.Assigned.ToString()))
            };
            if (IsFlexLicense(item))
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Dedicated machines you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Dedicated machines already created", item.DedicatedCloudPcCount.ToString())),
                    new Markup(PropertyInline("More dedicated machines you can create", item.AvailableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Shared pool Cloud PCs", item.SharedCloudPcCount.ToString())),
                    new Markup(PropertyInline("Total Flex Cloud PCs visible", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("Shared license units used", item.SharedUnitsUsed.ToString())),
                    new Markup(PropertyInline("Dedicated license units used", item.DedicatedUnitsUsed.ToString())),
                    new Markup(PropertyInline("Total license units used", item.LicenseUnitsUsed.ToString())),
                    new Markup(PropertyInline("License units left", item.LicenseUnitsLeft.ToString())),
                    new Markup(PropertyInline("Flex Cloud PCs that can have a user connected at once", item.ActiveSessionLimit.ToString()))
                ]);
            }
            else if (IsReserveLicense(item))
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Reserve Cloud PCs you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Reserve Cloud PCs already created", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("More Reserve Cloud PCs you can create", item.AvailableCloudPcCount.ToString()))
                ]);
            }
            else
            {
                detailRows.AddRange(
                [
                    new Markup(PropertyInline("Cloud PCs you can create", item.ProvisionableCloudPcCount.ToString())),
                    new Markup(PropertyInline("Cloud PCs already created", item.CloudPcCount.ToString())),
                    new Markup(PropertyInline("More Cloud PCs you can create", item.AvailableCloudPcCount.ToString()))
                ]);
            }
            var details = new Rows(detailRows);
            AnsiConsole.Write(new Panel(details).Header("License capacity").Border(BoxBorder.Rounded));
            AnsiConsole.WriteLine();
            AnsiConsole.Write(CreateLicensePlainEnglishPanel(item));

            if (IsFlexLicense(item))
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[#58a6ff]Flex rules[/]");
                AnsiConsole.MarkupLine($"[grey]Each Flex license unit can cover either 1 shared pool Cloud PC or up to 3 dedicated Cloud PCs. You have {item.LicenseUnitsLeft} license units left.[/]");
                AnsiConsole.WriteLine();
                var groupMembers = await LoadFlexPolicyGroupMembersAsync(item);
                AnsiConsole.Write(BuildCloudAppsPoolTable(item, groupMembers));
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildSharedPoolTable(item, groupMembers));
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildDedicatedMachineTable(item, groupMembers));
            }
            else
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(BuildLicenseCloudPcTable(item));
            }

            WaitForBack();
        }

        private static bool IsFlexLicense(LicenseOverviewItem item)
        {
            return item.Family.StartsWith("Flex", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsReserveLicense(LicenseOverviewItem item)
        {
            return item.Family.StartsWith("Reserve", StringComparison.OrdinalIgnoreCase);
        }

        private static Panel CreateLicensePlainEnglishPanel(LicenseOverviewItem item)
        {
            string[] lines;
            if (IsFlexLicense(item))
            {
                lines = new[]
                {
                    $"You bought {item.Purchased} Flex licenses for this size.",
                    $"{item.SharedCloudPcCount} shared pool or Cloud Apps Cloud PCs use {item.SharedUnitsUsed} license units.",
                    $"{item.DedicatedCloudPcCount} dedicated Cloud PCs use {item.DedicatedUnitsUsed} license units because dedicated capacity is counted in groups of 3.",
                    $"That uses {item.LicenseUnitsUsed} of {item.Purchased} license units, leaving {item.LicenseUnitsLeft}.",
                    $"You can create {item.AvailableCloudPcCount} more dedicated Cloud PCs. You can create {item.LicenseUnitsLeft} more shared pool Cloud PCs.",
                    $"Concurrency means {item.ActiveSessionLimit} total Flex Cloud PCs can have a user connected at the same time. It does not mean multiple users can connect to one Cloud PC."
                };
            }
            else if (IsReserveLicense(item))
            {
                lines =
                [
                    $"You bought {item.Purchased} Reserve licenses for this size.",
                    $"{item.CloudPcCount} Reserve Cloud PCs currently match this license size.",
                    $"{item.AvailableCloudPcCount} more Reserve Cloud PCs can be created before this license capacity is used.",
                    "The Cloud PC table below shows which user is assigned to each detected Reserve Cloud PC."
                ];
            }
            else
            {
                lines = new[]
                {
                    $"You bought {item.Purchased} licenses for this size.",
                    $"{item.CloudPcCount} Cloud PCs currently match this license size.",
                    $"{item.AvailableCloudPcCount} more Cloud PCs can be created before this license capacity is used."
                };
            }

            return new Panel(new Rows(lines.Select(line => new Markup($"[grey]{Markup.Escape(line)}[/]"))))
                .Header("What this means")
                .Border(BoxBorder.Rounded);
        }

        private static Table BuildLicenseCloudPcTable(LicenseOverviewItem item)
        {
            var table = new Table()
                .Border(TableBorder.Rounded)
                .AddColumn("Cloud PC")
                .AddColumn("Mode")
                .AddColumn("Assigned user")
                .AddColumn("Service plan")
                .AddColumn("Policy");

            foreach (var cloudPc in item.CloudPcs.OrderBy(pc => pc.ProvisioningPolicyName).ThenBy(pc => pc.Name, StringComparer.OrdinalIgnoreCase))
            {
                table.AddRow(
                    Markup.Escape(cloudPc.Name),
                    Markup.Escape(GetFlexCloudPcMode(cloudPc)),
                    Markup.Escape(cloudPc.UserPrincipalName ?? "-"),
                    Markup.Escape(cloudPc.ServicePlanName ?? "-"),
                    Markup.Escape(cloudPc.ProvisioningPolicyName ?? "-"));
            }

            if (item.CloudPcs.Count == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No matching Cloud PCs detected.[/]");
            }

            return table;
        }

        private async Task<IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>>> LoadFlexPolicyGroupMembersAsync(LicenseOverviewItem item)
        {
            var groupIds = item.FlexPolicies
                .SelectMany(policy => policy.AssignedGroupIds)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var resolved = await ConcurrencyHelper.MapWithConcurrencyAsync(groupIds, maxConcurrency: 5, async groupId =>
            {
                try
                {
                    return (groupId, members: (IReadOnlyList<GroupMemberSummary>)await _session.Graph.GetGroupMembersAsync(groupId));
                }
                catch (Exception)
                {
                    return (groupId, members: (IReadOnlyList<GroupMemberSummary>)Array.Empty<GroupMemberSummary>());
                }
            });

            return resolved.ToDictionary(pair => pair.groupId, pair => pair.members, StringComparer.OrdinalIgnoreCase);
        }

        private static Table BuildCloudAppsPoolTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Cloud Apps pool")
                .Border(TableBorder.Rounded)
                .AddColumn("Policy")
                .AddColumn("Group")
                .AddColumn("Cloud PC")
                .AddColumn("User")
                .AddColumn("UPN");

            var policies = item.FlexPolicies.Where(IsCloudAppsPolicy).ToArray();
            foreach (var policy in policies)
            {
                var cloudPcs = item.CloudPcs.Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase)).ToArray();
                var cloudPcList = cloudPcs.Length == 0 ? "-" : string.Join(", ", cloudPcs.Select(pc => pc.Name));
                foreach (var accessRow in GetPolicyAccessRows(policy, groupMembers))
                {
                    table.AddRow(
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(accessRow.GroupName),
                        Markup.Escape(cloudPcList),
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName));
                }
            }

            if (policies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No Cloud Apps pool detected.[/]", "[grey]-[/]");
            }

            return table;
        }

        private static Table BuildSharedPoolTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Shared pools")
                .Border(TableBorder.Rounded)
                .AddColumn("Pool")
                .AddColumn("Group")
                .AddColumn("Cloud PCs in pool")
                .AddColumn("User")
                .AddColumn("UPN");

            var policies = item.FlexPolicies.Where(policy => IsSharedFlexPolicy(policy) && !IsCloudAppsPolicy(policy) && !IsDedicatedFlexPolicy(policy)).ToArray();
            foreach (var policy in policies)
            {
                var cloudPcList = FormatCloudPcList(item.CloudPcs.Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase)));
                foreach (var accessRow in GetPolicyAccessRows(policy, groupMembers))
                {
                    table.AddRow(
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(accessRow.GroupName),
                        Markup.Escape(cloudPcList),
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName));
                }
            }

            if (policies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]No shared Flex pools detected.[/]", "[grey]-[/]", "[grey]-[/]");
            }

            return table;
        }

        private static Table BuildDedicatedMachineTable(LicenseOverviewItem item, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var table = new Table()
                .Title("Dedicated machines")
                .Border(TableBorder.Rounded)
                .AddColumn("User")
                .AddColumn("UPN")
                .AddColumn("Cloud PC")
                .AddColumn("Status")
                .AddColumn("Policy")
                .AddColumn("State");

            var dedicatedPolicies = item.FlexPolicies.Where(IsDedicatedFlexPolicy).ToArray();
            foreach (var policy in dedicatedPolicies)
            {
                var accessRows = GetPolicyAccessRows(policy, groupMembers)
                    .Where(row => row.UserPrincipalName != "-")
                    .OrderBy(row => row.UserName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var cloudPcsForPolicy = item.CloudPcs
                    .Where(pc => string.Equals(pc.ProvisioningPolicyId, policy.Id, StringComparison.OrdinalIgnoreCase) || string.Equals(pc.ProvisioningPolicyName, policy.DisplayName, StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var accessRow in accessRows)
                {
                    var cloudPc = cloudPcsForPolicy.FirstOrDefault(pc => string.Equals(pc.UserPrincipalName, accessRow.UserPrincipalName, StringComparison.OrdinalIgnoreCase));
                    table.AddRow(
                        Markup.Escape(accessRow.UserName),
                        Markup.Escape(accessRow.UserPrincipalName),
                        Markup.Escape(cloudPc?.Name ?? "-"),
                        Markup.Escape(cloudPc?.Status ?? "-"),
                        Markup.Escape(policy.DisplayName),
                        Markup.Escape(GetDedicatedUserState(cloudPc)));
                }
            }

            if (dedicatedPolicies.Length == 0)
            {
                table.AddRow("[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]-[/]", "[grey]No dedicated Flex policy detected.[/]", "[grey]-[/]");
            }

            return table;
        }

        private static string FormatCloudPcList(IEnumerable<CloudPcSummary> cloudPcs)
        {
            var names = cloudPcs.Select(pc => pc.Name).ToArray();
            return names.Length == 0 ? "-" : string.Join(", ", names);
        }

        private static IReadOnlyList<FlexAccessRow> GetPolicyAccessRows(ProvisioningPolicySummary policy, IReadOnlyDictionary<string, IReadOnlyList<GroupMemberSummary>> groupMembers)
        {
            var rows = new List<FlexAccessRow>();
            for (var index = 0; index < Math.Max(policy.AssignedGroupIds.Count, policy.AssignedGroupNames.Count); index++)
            {
                var groupId = index < policy.AssignedGroupIds.Count ? policy.AssignedGroupIds[index] : null;
                var groupName = index < policy.AssignedGroupNames.Count ? policy.AssignedGroupNames[index] : groupId ?? "-";
                var members = groupId is not null && groupMembers.TryGetValue(groupId, out var resolvedMembers)
                    ? resolvedMembers
                    : [];

                if (members.Count == 0)
                {
                    rows.Add(new FlexAccessRow(groupName, "No direct users found", "-"));
                    continue;
                }

                rows.AddRange(members.Select(member => new FlexAccessRow(groupName, member.Name, member.UserPrincipalName ?? "-")));
            }

            return rows.Count == 0 ? [new FlexAccessRow("-", "No direct users found", "-")] : rows;
        }

        private static string GetFlexCloudPcMode(CloudPcSummary cloudPc)
        {
            if (cloudPc.ProvisioningPolicyName?.Contains("Dedicated", StringComparison.OrdinalIgnoreCase) == true)
            {
                return "Dedicated";
            }

            return "Shared";
        }

        private static string GetDedicatedUserState(CloudPcSummary? cloudPc)
        {
            if (cloudPc is null)
            {
                return "Eligible, no Cloud PC yet";
            }

            var status = cloudPc.Status ?? string.Empty;
            return status.Contains("provisioning", StringComparison.OrdinalIgnoreCase) ||
                status.Contains("pending", StringComparison.OrdinalIgnoreCase)
                    ? "Provisioning"
                    : "Has Cloud PC";
        }

        private static bool IsCloudAppsPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Cloud-Apps", StringComparison.OrdinalIgnoreCase) ||
                policy.DisplayName.Contains("Cloud Apps", StringComparison.OrdinalIgnoreCase) ||
                policy.DisplayName.Contains("CloudApps", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSharedFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Shared", StringComparison.OrdinalIgnoreCase) ||
                (policy.ProvisioningType?.Contains("shared", StringComparison.OrdinalIgnoreCase) == true && !IsDedicatedFlexPolicy(policy));
        }

        private static bool IsDedicatedFlexPolicy(ProvisioningPolicySummary policy)
        {
            return policy.DisplayName.Contains("Dedicated", StringComparison.OrdinalIgnoreCase);
        }

    private sealed record LicenseOverviewItem(
        string Family,
        string SkuPartNumbers,
        int Purchased,
        int Assigned,
        int CloudPcCount,
        int DedicatedCloudPcCount,
        int SharedCloudPcCount,
        int ProvisionableCloudPcCount,
        int AvailableCloudPcCount,
        int ActiveSessionLimit,
        int DedicatedUnitsUsed,
        int SharedUnitsUsed,
        int LicenseUnitsUsed,
        int LicenseUnitsLeft,
        IReadOnlyList<CloudPcSummary> CloudPcs,
        IReadOnlyList<ProvisioningPolicySummary> FlexPolicies);

    private sealed record Windows365LicenseInfo(string Family, string PlanKey, string DisplayName);

    private sealed record FlexAccessRow(string GroupName, string UserName, string UserPrincipalName);
}
