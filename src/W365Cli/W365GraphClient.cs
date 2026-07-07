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
            "deviceManagement/virtualEndpoint/cloudPCs?$select=id,displayName,managedDeviceName,status,powerState,provisioningType,userPrincipalName,servicePlanName,managedDeviceId");

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

    public async Task RenameCloudPcAsync(string cloudPcId, string newDisplayName)
    {
        await PostJsonAsync(
            $"https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/rename",
            new
            {
                displayName = newDisplayName
            });
    }

    public async Task StartCloudPcAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/powerOn", new { });
    }

    public async Task EndCloudPcGracePeriodAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/endGracePeriod", new { });
    }

    public async Task ResetLocalAdminPasswordAsync(string managedDeviceId)
    {
        await PostJsonAsync($"deviceManagement/managedDevices('{Uri.EscapeDataString(managedDeviceId)}')/rotateLocalAdminPassword", new { });
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

    public async Task<IReadOnlyList<CloudPcSnapshot>> GetCloudPcSnapshotsAsync(CloudPcSummary cloudPc)
    {
        var escapedCloudPcId = Uri.EscapeDataString(cloudPc.Id);
        var select = Uri.EscapeDataString("id,cloudPcId,status,createdDateTime,lastRestoredDateTime,snapshotType,expirationDateTime,healthCheckStatus");
        var uri = $"deviceManagement/virtualEndpoint/cloudPCs/{escapedCloudPcId}/retrieveSnapshots?$select={select}";
        var page = await GetAsync<GraphPage<CloudPcSnapshotRaw>>(uri);

        return (page?.Value ?? [])
            .Select(snapshot => new CloudPcSnapshot
            {
                SnapshotId = snapshot.Id,
                CloudPcId = snapshot.CloudPcId ?? cloudPc.Id,
                Status = snapshot.Status,
                SnapshotType = snapshot.SnapshotType,
                CreatedDateTime = snapshot.CreatedDateTime,
                ExpirationDateTime = snapshot.ExpirationDateTime,
                LastRestoredDateTime = snapshot.LastRestoredDateTime,
                HealthCheckStatus = snapshot.HealthCheckStatus
            })
            .OrderByDescending(snapshot => snapshot.CreatedDateTime)
            .ToArray();
    }

    public async Task<IReadOnlyList<CloudPcServicePlan>> GetCloudPcServicePlansAsync()
    {
        var plans = await GetPagedAsync<CloudPcServicePlan>(
            "deviceManagement/virtualEndpoint/servicePlans");

        return plans
            .OrderBy(plan => plan.Type, StringComparer.OrdinalIgnoreCase)
            .ThenBy(plan => plan.VCpuCount)
            .ThenBy(plan => plan.RamGb)
            .ThenBy(plan => plan.StorageGb)
            .ThenBy(plan => plan.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public async Task ResizeCloudPcAsync(string cloudPcId, string targetServicePlanId)
    {
        await PostJsonAsync(
            $"https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/resize",
            new
            {
                targetServicePlanId
            });
    }

    public async Task CreateSnapshotAsync(string cloudPcId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/createSnapshot", new { });
    }

    public async Task RestoreSnapshotAsync(string cloudPcId, string snapshotId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/restore", new
        {
            cloudPcSnapshotId = snapshotId
        });
    }

    public async Task DeleteSnapshotAsync(string cloudPcId, string snapshotId)
    {
        var escapedSnapshotId = Uri.EscapeDataString(snapshotId);
        var attempts = new List<string>
        {
            $"deviceManagement/virtualEndpoint/snapshots/{escapedSnapshotId}"
        };

        if (!string.IsNullOrWhiteSpace(cloudPcId))
        {
            attempts.Add($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/snapshots/{escapedSnapshotId}");
        }

        var errors = new List<string>();
        foreach (var relativeUri in attempts)
        {
            using var request = new HttpRequestMessage(HttpMethod.Delete, relativeUri);
            await AuthorizeAsync(request);
            using var response = await _httpClient.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            var body = await response.Content.ReadAsStringAsync();
            errors.Add($"{relativeUri} returned {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
        }

        throw new HttpRequestException($"Snapshot delete failed. Attempted: {string.Join("; ", errors)}");
    }

    public async Task<IReadOnlyList<CloudPcRemoteActionResult>> GetCloudPcRemoteActionResultsAsync(CloudPcSummary cloudPc)
    {
        var escapedCloudPcId = Uri.EscapeDataString(cloudPc.Id);
        var uri = $"deviceManagement/virtualEndpoint/cloudPCs/{escapedCloudPcId}/retrieveCloudPCRemoteActionResults";
        var page = await GetAsync<GraphPage<CloudPcRemoteActionResultRaw>>(uri);

        return (page?.Value ?? [])
            .Select(result => new CloudPcRemoteActionResult
            {
                CloudPcId = cloudPc.Id,
                CloudPcName = cloudPc.Name,
                ActionName = result.ActionName,
                ActionState = result.ActionState,
                StartDateTime = result.StartDateTime?.ToLocalTime(),
                LastUpdatedDateTime = result.LastUpdatedDateTime?.ToLocalTime(),
                ManagedDeviceId = result.ManagedDeviceId,
                StatusCode = result.StatusDetail?.Code,
                StatusMessage = result.StatusDetail?.Message
            })
            .OrderByDescending(result => result.StartDateTime)
            .ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetOrganizationSettingsAsync()
    {
        return await GetJsonRowsAsync("deviceManagement/virtualEndpoint/organizationSettings", "id", "osVersion", "userAccountType", "windowsLanguage");
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetSettingProfilesAsync()
    {
        return await GetJsonRowsAsync("deviceManagement/virtualEndpoint/settingProfiles?$expand=assignments", "displayName", "profileType", "isAssigned", "lastModifiedDateTime");
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetUserSettingsAsync()
    {
        return await GetJsonRowsAsync("deviceManagement/virtualEndpoint/userSettings?$expand=assignments", "displayName", "selfServiceEnabled", "localAdminEnabled", "resetEnabled");
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetUsageRowsAsync()
    {
        var cloudPcs = await GetCloudPcsAsync();
        return cloudPcs
            .Select(pc => new GraphTableRow(
                pc.Name,
                JoinSummary(pc.Status, pc.PowerState, pc.ServicePlanName),
                new Dictionary<string, string>
                {
                    ["Cloud PC"] = pc.Name,
                    ["Status"] = pc.Status ?? "-",
                    ["Power state"] = pc.PowerState ?? "-",
                    ["Provisioning type"] = pc.ProvisioningType ?? "-",
                    ["User"] = pc.UserPrincipalName ?? "-",
                    ["Service plan"] = pc.ServicePlanName ?? "-",
                    ["Managed device"] = pc.ManagedDeviceName ?? "-",
                    ["Cloud PC ID"] = pc.Id,
                    ["Managed device ID"] = pc.ManagedDeviceId ?? "-"
                }))
            .ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetConnectivityHistoryAsync(CloudPcSummary cloudPc)
    {
        var rows = await GetJsonRowsAsync(
            $"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPc.Id)}/getCloudPcConnectivityHistory",
            "eventDateTime",
            "eventType",
            "eventName",
            "eventResult");

        return rows
            .Select(row =>
            {
                var fields = new Dictionary<string, string>(row.Fields)
                {
                    ["Cloud PC"] = cloudPc.Name
                };
                return row with { Fields = fields };
            })
            .ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetLaunchDetailRowsAsync()
    {
        var cloudPcs = await GetCloudPcsAsync();
        var rows = new List<GraphTableRow>();
        foreach (var cloudPc in cloudPcs)
        {
            if (string.IsNullOrWhiteSpace(cloudPc.UserPrincipalName))
            {
                rows.Add(new GraphTableRow(
                    cloudPc.Name,
                    "No user principal name",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Status"] = "Skipped",
                        ["Reason"] = "No user principal name"
                    }));
                continue;
            }

            try
            {
                var uri = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(cloudPc.UserPrincipalName)}/cloudPCs/{Uri.EscapeDataString(cloudPc.Id)}/retrieveCloudPcLaunchDetail";
                var item = await GetAsync<JsonElement>(uri);
                if (item.ValueKind == JsonValueKind.Object)
                {
                    var fields = FlattenJsonObject(item);
                    fields["Cloud PC"] = cloudPc.Name;
                    fields["User"] = cloudPc.UserPrincipalName;
                    rows.Add(ToTableRow(fields, "Cloud PC", "launchDetailStatus", "windows365SwitchCompatible"));
                }
            }
            catch (HttpRequestException ex)
            {
                rows.Add(new GraphTableRow(
                    cloudPc.Name,
                    "Launch details unavailable",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["User"] = cloudPc.UserPrincipalName,
                        ["Status"] = "Unavailable",
                        ["Error"] = ex.Message
                    }));
            }
        }

        return rows.OrderBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetCloudPcReportRowsAsync(string reportName, int top)
    {
        var definition = ResolveReportDefinition(reportName);
        var body = new Dictionary<string, object>
        {
            ["top"] = top
        };
        if (definition.IncludeReportName)
        {
            body["reportName"] = reportName;
        }

        var json = await PostJsonForStringAsync($"deviceManagement/virtualEndpoint/reports/{definition.Action}", body);
        using var document = JsonDocument.Parse(json);
        if (!document.RootElement.TryGetProperty("Schema", out var schema) ||
            !document.RootElement.TryGetProperty("Values", out var values) ||
            schema.ValueKind != JsonValueKind.Array ||
            values.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var columns = schema.EnumerateArray().Select(item => item.GetString() ?? "-").ToArray();
        var rows = new List<GraphTableRow>();
        foreach (var valueRow in values.EnumerateArray())
        {
            if (valueRow.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var valuesArray = valueRow.EnumerateArray().ToArray();
            for (var index = 0; index < Math.Min(columns.Length, valuesArray.Length); index++)
            {
                fields[columns[index]] = JsonToString(valuesArray[index]);
            }

            rows.Add(ToTableRow(fields, "CloudPcName", "ManagedDeviceName", "DisplayName", "SignInStatus", "Status", "Timestamp", "LastActiveTime"));
        }

        return rows;
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

    private async Task<IReadOnlyList<GraphTableRow>> GetJsonRowsAsync(string relativeUri, params string[] summaryFields)
    {
        var items = await GetPagedAsync<JsonElement>(relativeUri);
        return items
            .Where(item => item.ValueKind == JsonValueKind.Object)
            .Select(item => ToTableRow(FlattenJsonObject(item), summaryFields))
            .ToArray();
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

    private async Task<string> PostJsonForStringAsync(string relativeUri, object body)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, relativeUri)
        {
            Content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), System.Text.Encoding.UTF8, "application/json")
        };
        request.Headers.Add("Prefer", "include-unknown-enum-members");
        await AuthorizeAsync(request);
        using var response = await _httpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
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

    private static Dictionary<string, string> FlattenJsonObject(JsonElement item)
    {
        var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in item.EnumerateObject())
        {
            fields[property.Name] = JsonToString(property.Value);
        }

        return fields;
    }

    private static GraphTableRow ToTableRow(IReadOnlyDictionary<string, string> fields, params string[] summaryFields)
    {
        var title = GetFirst(fields, "displayName", "DisplayName", "Cloud PC", "CloudPcName", "ManagedDeviceName", "id") ?? "-";
        var summary = JoinSummary(summaryFields.Select(field => GetFirst(fields, field)).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray());
        return new GraphTableRow(title, string.IsNullOrWhiteSpace(summary) ? "-" : summary, fields);
    }

    private static string? GetFirst(IReadOnlyDictionary<string, string> fields, params string[] names)
    {
        foreach (var name in names)
        {
            if (fields.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static string JoinSummary(params string?[] values)
    {
        var parts = values.Where(value => !string.IsNullOrWhiteSpace(value) && value != "-").ToArray();
        return parts.Length == 0 ? "-" : string.Join(" | ", parts);
    }

    private static string JsonToString(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? "-",
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "-",
            JsonValueKind.Undefined => "-",
            _ => value.GetRawText()
        };
    }

    private static (string Action, bool IncludeReportName) ResolveReportDefinition(string reportName)
    {
        return reportName switch
        {
            "remoteConnectionHistoricalReports" => ("getRemoteConnectionHistoricalReports", false),
            "dailyAggregatedRemoteConnectionReports" => ("getDailyAggregatedRemoteConnectionReports", false),
            "totalAggregatedRemoteConnectionReports" => ("getTotalAggregatedRemoteConnectionReports", false),
            "frontlineLicenseUsageReport" => ("getFrontlineReport", true),
            "frontlineLicenseUsageRealTimeReport" => ("getFrontlineReport", true),
            "frontlineLicenseHourlyUsageReport" => ("getFrontlineReport", true),
            "frontlineRealtimeUserConnectionsReport" => ("getFrontlineReport", true),
            "inaccessibleCloudPcReports" => ("getInaccessibleCloudPcReports", true),
            "actionStatusReport" => ("getActionStatusReports", false),
            "performanceTrendReport" => ("retrieveCloudPcTenantMetricsReport", true),
            "regionalConnectionQualityTrendReport" => ("retrieveConnectionQualityReports", true),
            "cloudPcUsageCategoryReport" => ("retrieveCloudPcRecommendationReports", true),
            _ => throw new ArgumentOutOfRangeException(nameof(reportName), reportName, "Unknown Cloud PC report.")
        };
    }
}
