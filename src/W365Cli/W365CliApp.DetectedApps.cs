using Spectre.Console;

namespace W365Cli;

internal sealed partial class W365CliApp
{
    /// <summary>
    /// Tenant-wide Intune detected-apps inventory with a drill-down into which managed devices
    /// each app is installed on. Reuses the generic ShowGraphRowsAsync browser (filter/sort/CSV
    /// export already built in) rather than a bespoke screen like Cloud Apps' -- this is a plain
    /// read-only inventory report, not an entity with publish/unpublish actions.
    /// </summary>
    private async Task ShowDetectedAppsInventoryAsync()
    {
        if (!await EnsureConnectedAsync())
        {
            return;
        }

        await ShowGraphRowsAsync(
            "Detected Apps Inventory",
            async () => await _session.Graph.GetDetectedAppsAsync(),
            GetDetectedAppsHeader,
            FormatDetectedAppRow,
            ShowDetectedAppDevicesAsync);
    }

    private async Task ShowDetectedAppDevicesAsync(GraphTableRow app)
    {
        var appId = W365GraphClient.GetFirst(app.Fields, "id");
        if (string.IsNullOrWhiteSpace(appId))
        {
            TimedMessage("[yellow]This app row has no id to look up devices with.[/]");
            return;
        }

        await ShowGraphRowsAsync(
            $"Devices with {app.Title}",
            async () => await _session.Graph.GetDetectedAppDevicesAsync(appId),
            enterAction: OpenCloudPcFromDetectedAppDeviceAsync);
    }

    /// <summary>
    /// A device under a detected app is an Intune managedDevice, which may or may not be a
    /// Windows 365 Cloud PC. If it is, jump straight to the same Cloud PC details screen Cloud
    /// PCs > Browse uses -- matched on managedDeviceId == this device's own "id" first (the
    /// reliable FK Graph uses to link a cloudPC to its managedDevice record, so it matches even
    /// if the device/Cloud PC display names differ), then by name as a fallback. Falls back to
    /// the plain raw-field detail view (ShowGraphRowDetails) when no match is found.
    /// </summary>
    private async Task OpenCloudPcFromDetectedAppDeviceAsync(GraphTableRow device)
    {
        var cloudPcs = await LoadCloudPcsAsync();
        var cloudPc = ResolveCloudPcFromDetectedAppDevice(device, cloudPcs);
        if (cloudPc is null)
        {
            TimedMessage("[yellow]No matching Cloud PC found -- this device may not be a Cloud PC.[/]");
            ShowGraphRowDetails("Device details", device);
            return;
        }

        await ShowCloudPcDetailsAsync(cloudPc);
    }

    internal static CloudPcSummary? ResolveCloudPcFromDetectedAppDevice(GraphTableRow device, IReadOnlyList<CloudPcSummary> cloudPcs)
    {
        var deviceId = W365GraphClient.GetFirst(device.Fields, "id");
        if (!string.IsNullOrWhiteSpace(deviceId))
        {
            var idMatch = cloudPcs.FirstOrDefault(pc => string.Equals(pc.ManagedDeviceId, deviceId, StringComparison.OrdinalIgnoreCase));
            if (idMatch is not null)
            {
                return idMatch;
            }
        }

        var deviceName = GetOptionalField(device, "deviceName", "displayName", "managedDeviceName");
        return string.IsNullOrWhiteSpace(deviceName)
            ? null
            : cloudPcs.FirstOrDefault(pc =>
                string.Equals(pc.ManagedDeviceName, deviceName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pc.DisplayName, deviceName, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetDetectedAppsHeader()
    {
        var widths = GetDetectedAppsWidths();
        return Row(
            "App", widths.Name,
            "Version", widths.Version,
            "Publisher", widths.Publisher,
            "Platform", widths.Platform,
            "Devices", widths.Devices);
    }

    internal static (int Name, int Version, int Publisher, int Platform, int Devices) GetDetectedAppsWidths()
    {
        var available = Math.Max(70, Console.WindowWidth - 4);
        var name = Math.Max(24, Math.Min(40, available * 3 / 10));
        const int version = 16;
        var publisher = Math.Max(16, Math.Min(30, available / 6));
        const int platform = 10;
        const int devices = 9;
        return (name, version, publisher, platform, devices);
    }

    private static string FormatDetectedAppRow(GraphTableRow row)
    {
        var widths = GetDetectedAppsWidths();
        var version = W365GraphClient.GetFirst(row.Fields, "version") ?? "-";
        var publisher = W365GraphClient.GetFirst(row.Fields, "publisher") ?? "-";
        var platform = W365GraphClient.GetFirst(row.Fields, "platform") ?? "-";
        var deviceCount = W365GraphClient.GetFirst(row.Fields, "deviceCount") ?? "0";
        return Row(
            row.Title, widths.Name,
            version, widths.Version,
            publisher, widths.Publisher,
            platform, widths.Platform,
            deviceCount, widths.Devices);
    }

    /// <summary>
    /// Same app fields as the tenant-wide Detected Apps Inventory, minus the Devices column --
    /// Graph's deviceManagement/managedDevices('{id}')/detectedApps navigation property doesn't
    /// populate deviceCount (it comes back 0 for every row), unlike the top-level
    /// deviceManagement/detectedApps collection. Getting a real per-device-view count would mean
    /// a separate GetDetectedAppDevicesAsync call per app, i.e. exactly the N+1 throttling problem
    /// already solved for the tenant-wide inventory -- not worth it just for a count on a screen
    /// that's already scoped to one Cloud PC.
    /// </summary>
    private static string GetCloudPcApplicationsHeader()
    {
        var widths = GetDetectedAppsWidths();
        return Row(
            "App", widths.Name,
            "Version", widths.Version,
            "Publisher", widths.Publisher,
            "Platform", widths.Platform);
    }

    private static string FormatCloudPcApplicationRow(GraphTableRow row)
    {
        var widths = GetDetectedAppsWidths();
        var version = W365GraphClient.GetFirst(row.Fields, "version") ?? "-";
        var publisher = W365GraphClient.GetFirst(row.Fields, "publisher") ?? "-";
        var platform = W365GraphClient.GetFirst(row.Fields, "platform") ?? "-";
        return Row(
            row.Title, widths.Name,
            version, widths.Version,
            publisher, widths.Publisher,
            platform, widths.Platform);
    }
}
