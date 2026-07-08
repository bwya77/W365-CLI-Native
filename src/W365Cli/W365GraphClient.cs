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
            "deviceManagement/virtualEndpoint/cloudPCs?$select=id,displayName,managedDeviceName,status,powerState,provisioningType,userPrincipalName,servicePlanName,managedDeviceId,provisioningPolicyId,provisioningPolicyName");

        return items
            .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public async Task<IReadOnlyList<CloudPcSummary>> GetCloudPcsByProvisioningPolicyAsync(string provisioningPolicyId)
    {
        var filter = Uri.EscapeDataString($"servicePlanType eq 'enterprise' and provisioningPolicyId eq '{provisioningPolicyId}'");
        var select = Uri.EscapeDataString("id,displayName,managedDeviceName,status,powerState,provisioningType,userPrincipalName,servicePlanName,managedDeviceId,provisioningPolicyId,provisioningPolicyName");
        var items = await GetPagedAsync<CloudPcSummary>(
            $"deviceManagement/virtualEndpoint/cloudPCs?$filter={filter}&$select={select}");

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

    public async Task<IReadOnlyList<SubscribedSku>> GetSubscribedSkusAsync()
    {
        var select = Uri.EscapeDataString("skuId,skuPartNumber,prepaidUnits,consumedUnits,servicePlans");
        return await GetPagedAsync<SubscribedSku>($"https://graph.microsoft.com/v1.0/subscribedSkus?$select={select}");
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

    public async Task ReprovisionCloudPcAsync(string cloudPcId, string? osVersion = null, string? userAccountType = null)
    {
        var body = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(osVersion))
        {
            body["osVersion"] = osVersion;
        }
        if (!string.IsNullOrWhiteSpace(userAccountType))
        {
            body["userAccountType"] = userAccountType;
        }

        await PostJsonAsync($"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}/reprovision", body);
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

    public async Task<IReadOnlyList<GraphTableRow>> GetServicePlanRowsAsync()
    {
        var plans = await GetCloudPcServicePlansAsync();
        return plans
            .Select(plan => new GraphTableRow(
                plan.Name,
                JoinSummary(plan.Type, $"{plan.VCpuCount} vCPU", $"{plan.RamGb} GB RAM", $"{plan.StorageGb} GB storage"),
                new Dictionary<string, string>
                {
                    ["Name"] = plan.Name,
                    ["Type"] = plan.Type ?? "-",
                    ["vCPU"] = plan.VCpuCount?.ToString() ?? "-",
                    ["RAM"] = plan.RamGb is null ? "-" : $"{plan.RamGb} GB",
                    ["Storage"] = plan.StorageGb is null ? "-" : $"{plan.StorageGb} GB",
                    ["Profile"] = plan.UserProfileGb is null ? "-" : $"{plan.UserProfileGb} GB",
                    ["Service plan ID"] = plan.Id
                }))
            .ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetGalleryImageRowsAsync()
    {
        var select = Uri.EscapeDataString("id,displayName,offerDisplayName,skuDisplayName,publisherName,recommendedSku,status,sizeInGB,startDate,endDate,expirationDate,osVersionNumber");
        var rows = await GetJsonRowsAsync($"deviceManagement/virtualEndpoint/galleryImages?$select={select}", "status", "recommendedSku", "osVersionNumber");
        return rows.OrderBy(row => GetFirst(row.Fields, "status")).ThenBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetCustomImageRowsAsync()
    {
        var select = Uri.EscapeDataString("id,displayName,operatingSystem,osBuildNumber,version,status,expirationDate,osStatus,sourceImageResourceId,lastModifiedDateTime,statusDetails,errorCode,osVersionNumber,sizeInGB");
        var rows = await GetJsonRowsAsync($"deviceManagement/virtualEndpoint/deviceImages?$select={select}", "status", "operatingSystem", "osBuildNumber");
        return rows.OrderBy(row => GetFirst(row.Fields, "status")).ThenBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetSupportedRegionRowsAsync()
    {
        var select = Uri.EscapeDataString("id,displayName,regionStatus,supportedSolution,regionGroup,geographicLocationType");
        var rows = await GetJsonRowsAsync($"deviceManagement/virtualEndpoint/supportedRegions?$select={select}", "regionStatus", "supportedSolution", "regionGroup");
        return rows.OrderBy(row => GetFirst(row.Fields, "regionGroup")).ThenBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public async Task<IReadOnlyList<ProvisioningPolicySummary>> GetProvisioningPoliciesAsync()
    {
        var policies = await GetPagedAsync<JsonElement>("deviceManagement/virtualEndpoint/provisioningPolicies?$expand=assignments");
        var groupIds = policies
            .SelectMany(GetAssignmentGroupIds)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var groupNames = await ResolveGroupNamesAsync(groupIds);

        return policies
            .Select(policy => ToProvisioningPolicySummary(policy, groupNames))
            .OrderBy(policy => policy.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public async Task DeleteProvisioningPolicyAsync(string policyId)
    {
        using var request = new HttpRequestMessage(HttpMethod.Delete, $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}");
        await AuthorizeAsync(request);
        using var response = await _httpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();
    }

    public string ExportProvisioningPolicyJson(ProvisioningPolicySummary policy)
    {
        var export = new Dictionary<string, object?>
        {
            ["exportVersion"] = 1,
            ["exportedAt"] = DateTimeOffset.UtcNow.ToString("o"),
            ["sourceId"] = policy.Id,
            ["displayName"] = policy.DisplayName,
            ["createBody"] = BuildProvisioningPolicyCreateBody(policy, policy.DisplayName),
            ["assignments"] = BuildProvisioningPolicyAssignmentExports(policy)
        };

        return JsonSerializer.Serialize(export, new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            WriteIndented = true
        });
    }

    public async Task CreateProvisioningPolicyCopyAsync(ProvisioningPolicySummary policy, string displayName, bool assign)
    {
        var body = BuildProvisioningPolicyCreateBody(policy, displayName);
        var created = await PostJsonForElementAsync("deviceManagement/virtualEndpoint/provisioningPolicies", body);
        if (!assign)
        {
            return;
        }

        var createdId = GetString(created, "id");
        if (string.IsNullOrWhiteSpace(createdId))
        {
            throw new InvalidOperationException("Graph did not return the new provisioning policy id.");
        }

        var assignments = BuildProvisioningPolicyAssignmentsForCreate(policy, createdId);
        if (assignments.Count == 0)
        {
            return;
        }

        await PostJsonAsync($"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(createdId)}/assign", new
        {
            assignments
        });
    }

    public async Task ReprovisionCloudPcsByPolicyAsync(string policyId, string? osVersion, string? userAccountType, IReadOnlyList<string> exclusions)
    {
        var cloudPcs = await GetCloudPcsByProvisioningPolicyAsync(policyId);
        var excludeSet = new HashSet<string>(exclusions.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value.Trim()), StringComparer.OrdinalIgnoreCase);
        foreach (var cloudPc in cloudPcs)
        {
            var matchValues = new[]
            {
                cloudPc.Id,
                cloudPc.Name,
                cloudPc.ManagedDeviceId,
                cloudPc.UserPrincipalName
            }.Where(value => !string.IsNullOrWhiteSpace(value));

            if (matchValues.Any(value => excludeSet.Contains(value!)))
            {
                continue;
            }

            await ReprovisionCloudPcAsync(cloudPc.Id, osVersion, userAccountType);
        }
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
                    "Status: Skipped | Reason: No user principal name",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Cloud PC ID"] = cloudPc.Id,
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
                    fields["Cloud PC ID"] = cloudPc.Id;
                    fields["User"] = cloudPc.UserPrincipalName;
                    fields["Status"] = "Available";
                    var status = "Available";
                    var switchCompatible = FormatBoolean(GetFirst(fields, "windows365SwitchCompatible"));
                    rows.Add(new GraphTableRow(
                        cloudPc.Name,
                        JoinSummary($"Status: {status}", $"Switch compatible: {switchCompatible}"),
                        fields));
                }
            }
            catch (HttpRequestException ex)
            {
                var reason = FormatLaunchDetailError(ex);
                rows.Add(new GraphTableRow(
                    cloudPc.Name,
                    $"Status: Unavailable | Reason: {reason}",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Cloud PC ID"] = cloudPc.Id,
                        ["User"] = cloudPc.UserPrincipalName,
                        ["Status"] = "Unavailable",
                        ["Reason"] = reason
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

        var columns = schema.EnumerateArray().Select(GetReportColumnName).ToArray();
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

    private async Task<JsonElement> PostJsonForElementAsync(string relativeUri, object body)
    {
        var json = await PostJsonForStringAsync(relativeUri, body);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
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

    private static ProvisioningPolicySummary ToProvisioningPolicySummary(JsonElement policy, IReadOnlyDictionary<string, string> groupNames)
    {
        var groupIds = GetAssignmentGroupIds(policy).ToArray();
        var domainJoinTypes = policy.TryGetProperty("domainJoinConfigurations", out var joins) && joins.ValueKind == JsonValueKind.Array
            ? string.Join(",", joins.EnumerateArray().Select(join => GetString(join, "domainJoinType")).Where(value => !string.IsNullOrWhiteSpace(value)))
            : null;

        return new ProvisioningPolicySummary
        {
            Id = GetString(policy, "id") ?? string.Empty,
            DisplayName = GetString(policy, "displayName") ?? GetString(policy, "id") ?? "-",
            Description = GetString(policy, "description"),
            ProvisioningType = GetString(policy, "provisioningType"),
            ImageDisplayName = GetString(policy, "imageDisplayName"),
            ImageType = GetString(policy, "imageType"),
            DomainJoinTypes = domainJoinTypes,
            EnableSingleSignOn = GetBool(policy, "enableSingleSignOn"),
            LocalAdminEnabled = GetBool(policy, "localAdminEnabled"),
            CloudPcNamingTemplate = GetString(policy, "cloudPcNamingTemplate"),
            CloudPcGroupDisplayName = GetString(policy, "cloudPcGroupDisplayName"),
            ManagedBy = GetString(policy, "managedBy"),
            GracePeriodInHours = GetInt(policy, "gracePeriodInHours"),
            AssignedGroupIds = groupIds,
            AssignedGroupNames = groupIds.Select(groupId => groupNames.TryGetValue(groupId, out var name) ? name : groupId).ToArray(),
            Raw = policy.Clone()
        };
    }

    private static IEnumerable<string> GetAssignmentGroupIds(JsonElement policy)
    {
        if (!policy.TryGetProperty("assignments", out var assignments) || assignments.ValueKind != JsonValueKind.Array)
        {
            yield break;
        }

        foreach (var assignment in assignments.EnumerateArray())
        {
            if (assignment.TryGetProperty("target", out var target))
            {
                var groupId = GetString(target, "groupId");
                if (!string.IsNullOrWhiteSpace(groupId))
                {
                    yield return groupId;
                }
            }
        }
    }

    private async Task<IReadOnlyDictionary<string, string>> ResolveGroupNamesAsync(IReadOnlyList<string> groupIds)
    {
        var output = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var groupId in groupIds)
        {
            try
            {
                var group = await GetAsync<JsonElement>($"groups/{Uri.EscapeDataString(groupId)}?$select=id,displayName");
                var name = group.ValueKind == JsonValueKind.Object ? GetString(group, "displayName") : null;
                output[groupId] = string.IsNullOrWhiteSpace(name) ? groupId : name;
            }
            catch (HttpRequestException)
            {
                output[groupId] = groupId;
            }
        }

        return output;
    }

    private static Dictionary<string, object?> BuildProvisioningPolicyCreateBody(ProvisioningPolicySummary policy, string displayName)
    {
        var createKeys = new[]
        {
            "@odata.type",
            "autopatch",
            "cloudPcNamingTemplate",
            "description",
            "displayName",
            "domainJoinConfigurations",
            "enableSingleSignOn",
            "imageDisplayName",
            "imageId",
            "imageType",
            "localAdminEnabled",
            "managedBy",
            "microsoftManagedDesktop",
            "provisioningType",
            "userExperienceType",
            "userSettingsPersistenceConfiguration",
            "windowsSetting"
        };

        var body = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in createKeys)
        {
            if (policy.Raw.TryGetProperty(key, out var value) && value.ValueKind is not JsonValueKind.Null and not JsonValueKind.Undefined)
            {
                body[key] = value.Clone();
            }
        }

        body["@odata.type"] = "#microsoft.graph.cloudPcProvisioningPolicy";
        body["displayName"] = displayName;
        if (!body.ContainsKey("description"))
        {
            body["description"] = string.Empty;
        }

        return body;
    }

    private static IReadOnlyList<object> BuildProvisioningPolicyAssignmentExports(ProvisioningPolicySummary policy)
    {
        return BuildProvisioningPolicyAssignments(policy, null, includeSourceId: true);
    }

    private static IReadOnlyList<object> BuildProvisioningPolicyAssignmentsForCreate(ProvisioningPolicySummary policy, string createdPolicyId)
    {
        return BuildProvisioningPolicyAssignments(policy, createdPolicyId, includeSourceId: false);
    }

    private static IReadOnlyList<object> BuildProvisioningPolicyAssignments(ProvisioningPolicySummary policy, string? createdPolicyId, bool includeSourceId)
    {
        if (!policy.Raw.TryGetProperty("assignments", out var assignments) || assignments.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var output = new List<object>();
        foreach (var assignment in assignments.EnumerateArray())
        {
            if (!assignment.TryGetProperty("target", out var target))
            {
                continue;
            }

            var groupId = GetString(target, "groupId");
            if (string.IsNullOrWhiteSpace(groupId))
            {
                continue;
            }

            var targetBody = new Dictionary<string, object?>
            {
                ["@odata.type"] = GetString(target, "@odata.type") ?? "microsoft.graph.cloudPcManagementGroupAssignmentTarget",
                ["groupId"] = groupId
            };

            AddJsonValueIfPresent(targetBody, target, "servicePlanId");
            AddJsonValueIfPresent(targetBody, target, "allotmentLicensesCount");
            AddJsonValueIfPresent(targetBody, target, "allotmentDisplayName");

            var assignmentBody = new Dictionary<string, object?>
            {
                ["target"] = targetBody
            };

            if (includeSourceId)
            {
                assignmentBody["sourceId"] = GetString(assignment, "id");
            }
            else if (string.Equals(policy.ProvisioningType, "dedicated", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(createdPolicyId))
            {
                assignmentBody["id"] = $"{createdPolicyId}_{groupId}";
            }

            output.Add(assignmentBody);
        }

        return output;
    }

    private static void AddJsonValueIfPresent(IDictionary<string, object?> target, JsonElement source, string property)
    {
        if (source.TryGetProperty(property, out var value) && value.ValueKind is not JsonValueKind.Null and not JsonValueKind.Undefined)
        {
            target[property] = value.Clone();
        }
    }

    private static string? GetString(JsonElement item, string property)
    {
        if (!item.TryGetProperty(property, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => null
        };
    }

    private static bool? GetBool(JsonElement item, string property)
    {
        if (!item.TryGetProperty(property, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(value.GetString(), out var parsed) => parsed,
            _ => null
        };
    }

    private static int? GetInt(JsonElement item, string property)
    {
        if (!item.TryGetProperty(property, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.Number when value.TryGetInt32(out var parsed) => parsed,
            JsonValueKind.String when int.TryParse(value.GetString(), out var parsed) => parsed,
            _ => null
        };
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

    private static string GetReportColumnName(JsonElement schemaItem)
    {
        if (schemaItem.ValueKind == JsonValueKind.String)
        {
            return schemaItem.GetString() ?? "-";
        }

        if (schemaItem.ValueKind == JsonValueKind.Object)
        {
            foreach (var name in new[] { "Column", "column", "Name", "name" })
            {
                if (schemaItem.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String)
                {
                    var value = property.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return value;
                    }
                }
            }
        }

        return "-";
    }

    private static string FormatBoolean(string? value)
    {
        return value?.ToLowerInvariant() switch
        {
            "true" => "Yes",
            "false" => "No",
            null or "" or "-" => "Unknown",
            _ => value
        };
    }

    private static string FormatLaunchDetailError(HttpRequestException ex)
    {
        var message = ex.Message;
        if (message.Contains("404", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Not Found", StringComparison.OrdinalIgnoreCase))
        {
            return "Launch details not found for this user";
        }

        if (message.Contains("403", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Forbidden", StringComparison.OrdinalIgnoreCase))
        {
            return "Access denied";
        }

        if (message.Contains("401", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase))
        {
            return "Authentication required";
        }

        return "Launch details unavailable";
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
