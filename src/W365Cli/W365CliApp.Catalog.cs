using Spectre.Console;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace W365Cli;

internal sealed partial class W365CliApp
{

    private async Task ShowTenantSettingsAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        while (true)
        {
            var choice = PromptChoice(
                () => { },
                "[#58a6ff]Tenant settings[/]",
                ["Organization settings", "Setting profiles", "User settings", "Back"],
                "Back");

            switch (choice)
            {
                case "Organization settings":
                    await ShowGraphRowsAsync(
                        "Windows 365 organization settings",
                        async () => await _session.Graph.GetOrganizationSettingsAsync(),
                        GetOrganizationSettingsHeader,
                        FormatOrganizationSettingRow);
                    break;
                case "Setting profiles":
                    await ShowGraphRowsAsync(
                        "Windows 365 setting profiles",
                        async () => await _session.Graph.GetSettingProfilesAsync(),
                        GetSettingProfilesHeader,
                        FormatSettingProfileRow);
                    break;
                case "User settings":
                    await ShowGraphRowsAsync(
                        "Windows 365 user settings",
                        async () => await _session.Graph.GetUserSettingsAsync(),
                        GetUserSettingsHeader,
                        FormatUserSettingRow);
                    break;
                case "Back":
                    return;
            }
        }
    }

    private async Task ShowCatalogAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        var choices = new[] { "Service plans", "Gallery images", "Custom images", "Supported regions", "Back" };
        var selectedIndex = 0;
        while (true)
        {
            AnsiConsole.Clear();
            RenderBreadcrumb("Catalog");
            AnsiConsole.MarkupLine("[#58a6ff]Catalog[/]");
            AnsiConsole.MarkupLine("[grey]Plans, images, and regions used by Windows 365.[/]");
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
                        case "Service plans":
                            await ShowGraphRowsAsync("Windows 365 service plans", _session.Graph.GetServicePlanRowsAsync, GetServicePlansHeader, FormatServicePlanRow);
                            break;
                        case "Gallery images":
                            await ShowGraphRowsAsync("Windows 365 gallery images", _session.Graph.GetGalleryImageRowsAsync, GetGalleryImagesHeader, FormatGalleryImageRow);
                            break;
                        case "Custom images":
                            await ShowGraphRowsAsync("Windows 365 custom images", _session.Graph.GetCustomImageRowsAsync, GetCustomImagesHeader, FormatCustomImageRow);
                            break;
                        case "Supported regions":
                            await ShowGraphRowsAsync("Windows 365 supported regions", _session.Graph.GetSupportedRegionRowsAsync, GetSupportedRegionsHeader, FormatSupportedRegionRow);
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
                    if (key.KeyChar is 'p' or 'P')
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

    private static string GetServicePlansHeader()
    {
        return Row("Name", 44, "Type", 12, "vCPU", 6, "RAM", 8, "Storage", 10, "Profile", 10);
    }

    private static string FormatServicePlanRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "Name"), 44,
            GetField(row, "Type"), 12,
            GetField(row, "vCPU"), 6,
            GetField(row, "RAM"), 8,
            GetField(row, "Storage"), 10,
            GetField(row, "Profile"), 10);
    }

    private static string GetGalleryImagesHeader()
    {
        var widths = GetCatalogImageWidths();
        return Row("Name", widths.Name, "Status", widths.Status, "Recommended SKU", widths.Sku, "Size", widths.Size, "OS version", widths.Os);
    }

    private static string FormatGalleryImageRow(GraphTableRow row)
    {
        var widths = GetCatalogImageWidths();
        return Row(
            GetField(row, "displayName"), widths.Name,
            GetField(row, "status"), widths.Status,
            GetField(row, "recommendedSku"), widths.Sku,
            FormatCatalogGb(GetField(row, "sizeInGB")), widths.Size,
            GetField(row, "osVersionNumber"), widths.Os);
    }

    private static string GetCustomImagesHeader()
    {
        var widths = GetCatalogImageWidths();
        return Row("Name", widths.Name, "Status", widths.Status, "OS", widths.Sku, "Size", widths.Size, "Modified", widths.Os);
    }

    private static string FormatCustomImageRow(GraphTableRow row)
    {
        var widths = GetCatalogImageWidths();
        return Row(
            GetField(row, "displayName"), widths.Name,
            GetField(row, "status"), widths.Status,
            GetField(row, "operatingSystem"), widths.Sku,
            FormatCatalogGb(GetField(row, "sizeInGB")), widths.Size,
            GetField(row, "lastModifiedDateTime"), widths.Os);
    }

    private static (int Name, int Status, int Sku, int Size, int Os) GetCatalogImageWidths()
    {
        var available = Math.Max(92, Console.WindowWidth - 4);
        const int status = 14;
        const int size = 8;
        var remaining = Math.Max(50, available - status - size - 4);
        var name = Math.Clamp((int)(remaining * 0.42), 30, 44);
        var sku = Math.Clamp((int)(remaining * 0.34), 18, 30);
        var os = Math.Max(16, remaining - name - sku);
        return (name, status, sku, size, os);
    }

    private static string GetSupportedRegionsHeader()
    {
        return Row("Name", 34, "Status", 12, "Solution", 16, "Group", 20, "Geo", 20);
    }

    private static string FormatSupportedRegionRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 34,
            GetField(row, "regionStatus"), 12,
            GetField(row, "supportedSolution"), 16,
            GetField(row, "regionGroup"), 20,
            GetField(row, "geographicLocationType"), 20);
    }

    private static string FormatCatalogGb(string value)
    {
        return value == "-" ? "-" : $"{value} GB";
    }

    private static string GetOrganizationSettingsHeader()
    {
        return Row("OS", 14, "User", 18, "MEM auto", 10, "SSO", 8, "Language", 14);
    }

    private static string FormatOrganizationSettingRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "osVersion"), 14,
            GetField(row, "userAccountType"), 18,
            FormatBooleanCell(GetField(row, "memAutoEnrollEnabled")), 10,
            FormatBooleanCell(GetField(row, "singleSignOnEnabled")), 8,
            GetField(row, "windowsLanguage"), 14);
    }

    private static string GetSettingProfilesHeader()
    {
        return Row("Name", 42, "Type", 18, "Assigned", 10, "Priority", 10, "Modified", 20);
    }

    private static string FormatSettingProfileRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 42,
            GetField(row, "profileType"), 18,
            FormatBooleanCell(GetField(row, "isAssigned")), 10,
            GetNestedField(row, "priorityMetaData", "priority"), 10,
            FormatDateCell(GetField(row, "lastModifiedDateTime")), 20);
    }

    private static string GetUserSettingsHeader()
    {
        return Row("Name", 38, "Self svc", 10, "Admin", 8, "Reset", 8, "Restore", 9, "DR", 8);
    }

    private static string FormatUserSettingRow(GraphTableRow row)
    {
        return Row(
            GetField(row, "displayName"), 38,
            FormatBooleanCell(GetField(row, "selfServiceEnabled")), 10,
            FormatBooleanCell(GetField(row, "localAdminEnabled")), 8,
            FormatBooleanCell(GetField(row, "resetEnabled")), 8,
            FormatBooleanCell(GetNestedField(row, "restorePointSetting", "userRestoreEnabled")), 9,
            FormatBooleanCell(GetNestedField(row, "crossRegionDisasterRecoverySetting", "crossRegionDisasterRecoveryEnabled")), 8);
    }
}
