using System.Net.Http.Headers;
using System.Text.Json;

namespace W365Cli;

internal sealed class W365GraphClient
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly Func<Task<string>>? _accessTokenProvider;
    private readonly HttpClient _httpClient;

    public static W365GraphClient NotConnected { get; } = new(null);

    public W365GraphClient(Func<Task<string>>? accessTokenProvider)
    {
        _accessTokenProvider = accessTokenProvider;
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri("https://graph.microsoft.com/beta/")
        };
    }

    public async Task<IReadOnlyList<CloudPcSummary>> GetCloudPcsAsync()
    {
        var items = await GetPagedAsync<CloudPcSummary>(
            "deviceManagement/virtualEndpoint/cloudPCs?$select=id,displayName,managedDeviceName,status,provisioningType,userPrincipalName,servicePlanName,managedDeviceId");

        return items
            .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public async Task<IReadOnlyList<CloudAppSummary>> GetCloudAppsAsync()
    {
        var items = await GetPagedAsync<CloudAppSummary>(
            "deviceManagement/virtualEndpoint/cloudApps?$top=100&$orderBy=lastPublishedDateTime desc&$count=true&$select=*",
            includeConsistencyLevel: true);

        return items
            .OrderByDescending(item => item.LastPublishedDateTime)
            .ToArray();
    }

    public async Task<OrganizationSummary?> GetOrganizationAsync()
    {
        var items = await GetPagedAsync<OrganizationSummary>(
            "organization?$select=id,displayName");

        return items.FirstOrDefault();
    }

    public async Task PublishCloudAppAsync(string cloudAppId)
    {
        await PostJsonAsync("deviceManagement/virtualEndpoint/cloudApps/publish", new
        {
            cloudAppIds = new[] { cloudAppId }
        });
    }

    public async Task UnpublishCloudAppAsync(string cloudAppId)
    {
        await PostJsonAsync("deviceManagement/virtualEndpoint/cloudApps/unpublish", new
        {
            cloudAppIds = new[] { cloudAppId }
        });
    }

    public async Task RestartCloudPcAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/reboot", new { });
    }

    public async Task StartCloudPcAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/powerOn", new { });
    }

    public async Task ReprovisionCloudPcAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/reprovision", new { });
    }

    public async Task SyncManagedDeviceAsync(string managedDeviceId)
    {
        await PostJsonAsync($"deviceManagement/managedDevices/{Uri.EscapeDataString(managedDeviceId)}/syncDevice", new { });
    }

    public async Task<IReadOnlyList<CloudPcDiskSpace>> GetCloudPcDiskSpacesAsync(IReadOnlyList<CloudPcSummary>? cloudPcs = null)
    {
        cloudPcs ??= await GetCloudPcsAsync();
        var results = new List<CloudPcDiskSpace>();

        foreach (var cloudPc in cloudPcs)
        {
            ManagedDeviceDiskInfo? managedDevice = null;
            if (!string.IsNullOrWhiteSpace(cloudPc.ManagedDeviceId))
            {
                var escapedManagedDeviceId = Uri.EscapeDataString(cloudPc.ManagedDeviceId);
                managedDevice = await GetAsync<ManagedDeviceDiskInfo>(
                    $"deviceManagement/managedDevices/{escapedManagedDeviceId}?$select=id,deviceName,totalStorageSpaceInBytes,freeStorageSpaceInBytes,lastSyncDateTime");
            }

            var totalGb = ToGb(managedDevice?.TotalStorageSpaceInBytes);
            var freeGb = ToGb(managedDevice?.FreeStorageSpaceInBytes);
            double? usedGb = totalGb is not null && freeGb is not null
                ? Math.Round(totalGb.Value - freeGb.Value, 2)
                : null;
            double? percentFree = totalGb is > 0 && freeGb is not null
                ? Math.Round((freeGb.Value / totalGb.Value) * 100, 1)
                : null;

            results.Add(new CloudPcDiskSpace
            {
                CloudPcId = cloudPc.Id,
                CloudPcName = cloudPc.Name,
                AssignedUserUpn = cloudPc.UserPrincipalName,
                ManagedDeviceId = cloudPc.ManagedDeviceId,
                ManagedDeviceName = managedDevice?.DeviceName ?? cloudPc.ManagedDeviceName,
                TotalStorageGb = totalGb,
                FreeStorageGb = freeGb,
                UsedStorageGb = usedGb,
                PercentFree = percentFree,
                LastSyncDateTime = managedDevice?.LastSyncDateTime
            });
        }

        return results
            .OrderBy(item => item.PercentFree ?? double.MaxValue)
            .ThenBy(item => item.CloudPcName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private async Task<List<T>> GetPagedAsync<T>(string relativeUri, bool includeConsistencyLevel = false)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        var output = new List<T>();
        var next = relativeUri;

        while (!string.IsNullOrWhiteSpace(next))
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, next);
            if (includeConsistencyLevel)
            {
                request.Headers.Add("ConsistencyLevel", "eventual");
            }

            await AuthorizeAsync(request);
            using var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();

            await using var stream = await response.Content.ReadAsStreamAsync();
            var page = await JsonSerializer.DeserializeAsync<GraphPage<T>>(stream, JsonOptions);
            if (page?.Value is not null)
            {
                output.AddRange(page.Value);
            }

            next = page?.NextLink;
        }

        return output;
    }

    private async Task<T?> GetAsync<T>(string relativeUri)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, relativeUri);
        await AuthorizeAsync(request);
        using var response = await _httpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync();
        return await JsonSerializer.DeserializeAsync<T>(stream, JsonOptions);
    }

    private async Task PostJsonAsync(string relativeUri, object body)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, relativeUri)
        {
            Content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), System.Text.Encoding.UTF8, "application/json")
        };
        await AuthorizeAsync(request);
        using var response = await _httpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();
    }

    private async Task AuthorizeAsync(HttpRequestMessage request)
    {
        var token = await _accessTokenProvider!();
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    }

    private static double? ToGb(long? bytes)
    {
        return bytes is null ? null : Math.Round(bytes.Value / 1024d / 1024d / 1024d, 2);
    }
}
