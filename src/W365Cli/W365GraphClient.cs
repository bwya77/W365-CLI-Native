using System.Net;
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
            "deviceManagement/virtualEndpoint/cloudPCs?$select=id,displayName,managedDeviceName,status,powerState,provisioningType,userPrincipalName,servicePlanName,managedDeviceId,provisioningPolicyId,provisioningPolicyName,sharedDeviceDetail");

        return items
            .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public async Task<IReadOnlyList<CloudPcSummary>> GetCloudPcsByProvisioningPolicyAsync(string provisioningPolicyId)
    {
        var filter = Uri.EscapeDataString($"provisioningPolicyId eq '{provisioningPolicyId}' and servicePlanType eq 'enterprise'");
        var select = Uri.EscapeDataString("id,displayName,managedDeviceName,status,powerState,provisioningType,userPrincipalName,servicePlanName,managedDeviceId,provisioningPolicyId,provisioningPolicyName,sharedDeviceDetail,connectivityResult");
        var items = await GetPagedAsync<CloudPcSummary>(
            $"deviceManagement/virtualEndpoint/cloudPCs?$filter={filter}&$select={select}&$orderBy=lastModifiedDateTime desc&$count=true",
            includeConsistencyLevel: true,
            includeUnknownEnumMembers: true);

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

    public async Task<IReadOnlyList<GroupMemberSummary>> GetGroupMembersAsync(string groupId)
    {
        var select = Uri.EscapeDataString("id,displayName,userPrincipalName");
        return await GetPagedAsync<GroupMemberSummary>($"https://graph.microsoft.com/v1.0/groups/{Uri.EscapeDataString(groupId)}/members/microsoft.graph.user?$select={select}");
    }

    /// <summary>
    /// Fast member-count lookup for a group, without pulling every member's details — matches the
    /// same $count endpoint the Windows 365 admin portal calls right after picking a group in the
    /// Flex assignment wizard, so the CLI can show the same "your group has N members" context.
    /// Returns null if the count can't be determined (best-effort only; never blocks the wizard).
    /// </summary>
    public async Task<int?> GetGroupMemberCountAsync(string groupId)
    {
        try
        {
            var countText = await GetRawStringAsync(
                $"https://graph.microsoft.com/v1.0/groups/{Uri.EscapeDataString(groupId)}/members/microsoft.graph.user/$count",
                includeConsistencyLevel: true);
            return int.TryParse(countText, out var count) ? count : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Searches Entra groups by display name prefix, for picking which group a provisioning
    /// policy (or other assignment) targets instead of requiring the user to paste a raw object
    /// ID. Uses advanced query ($count + ConsistencyLevel eventual) to match SearchUsersAsync's
    /// pattern, since startswith() filters need it for reliable paging/counting.
    /// </summary>
    public async Task<IReadOnlyList<EntraGroupSummary>> SearchGroupsAsync(string query)
    {
        var escaped = query.Replace("'", "''", StringComparison.Ordinal);
        var filter = Uri.EscapeDataString($"startswith(displayName,'{escaped}') or startswith(mailNickname,'{escaped}')");
        var select = Uri.EscapeDataString("id,displayName,mailNickname");
        return await GetPagedAsync<EntraGroupSummary>(
            $"https://graph.microsoft.com/v1.0/groups?$filter={filter}&$select={select}&$top=25&$count=true",
            includeConsistencyLevel: true);
    }

    /// <summary>
    /// Searches directory users by display name or UPN prefix, for picking someone to add to a
    /// provisioning policy's assigned Entra group. Uses advanced query ($count + ConsistencyLevel
    /// eventual) since the filter combines startswith() across two properties with "or".
    /// </summary>
    public async Task<IReadOnlyList<GroupMemberSummary>> SearchUsersAsync(string query)
    {
        var escaped = query.Replace("'", "''", StringComparison.Ordinal);
        var filter = Uri.EscapeDataString($"startswith(displayName,'{escaped}') or startswith(userPrincipalName,'{escaped}') or startswith(mail,'{escaped}')");
        var select = Uri.EscapeDataString("id,displayName,userPrincipalName");
        return await GetPagedAsync<GroupMemberSummary>(
            $"https://graph.microsoft.com/v1.0/users?$filter={filter}&$select={select}&$top=25&$count=true",
            includeConsistencyLevel: true);
    }

    /// <summary>
    /// Adds a user to an Entra group (e.g. a provisioning policy's assigned group) via the
    /// $ref navigation-property POST.
    /// </summary>
    public async Task AddGroupMemberAsync(string groupId, string userId)
    {
        await PostJsonAsync(
            $"https://graph.microsoft.com/v1.0/groups/{Uri.EscapeDataString(groupId)}/members/$ref",
            new Dictionary<string, object>
            {
                ["@odata.id"] = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
            });
    }

    /// <summary>
    /// Removes a user from an Entra group via the $ref navigation-property DELETE.
    /// </summary>
    public async Task RemoveGroupMemberAsync(string groupId, string userId)
    {
        using var response = await SendWithRetryAsync(() => new HttpRequestMessage(
            HttpMethod.Delete,
            $"https://graph.microsoft.com/v1.0/groups/{Uri.EscapeDataString(groupId)}/members/{Uri.EscapeDataString(userId)}/$ref"));
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
        }
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

    /// <summary>
    /// Fetches a Cloud PC's <c>statusDetail</c> (why it's provisionedWithWarnings/
    /// provisionedWithErrors/failed) — the same data the Intune portal's "View more information"
    /// panel is built from. Returns null if the Cloud PC has no status detail (e.g. healthy).
    /// </summary>
    public async Task<CloudPcStatusDetail?> GetCloudPcStatusDetailAsync(string cloudPcId)
    {
        var select = Uri.EscapeDataString("id,statusDetail");
        var cloudPc = await GetAsync<JsonElement>(
            $"deviceManagement/virtualEndpoint/cloudPCs/{Uri.EscapeDataString(cloudPcId)}?$select={select}");

        if (!cloudPc.TryGetProperty("statusDetail", out var detail) || detail.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        var code = GetString(detail, "code");
        var message = GetString(detail, "message");

        var additionalInfo = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (detail.TryGetProperty("additionalInformation", out var info) && info.ValueKind == JsonValueKind.Array)
        {
            foreach (var entry in info.EnumerateArray())
            {
                var name = GetString(entry, "name");
                var value = GetString(entry, "value");
                if (!string.IsNullOrWhiteSpace(name))
                {
                    additionalInfo[name] = value ?? string.Empty;
                }
            }
        }

        if (string.IsNullOrWhiteSpace(code) && string.IsNullOrWhiteSpace(message) && additionalInfo.Count == 0)
        {
            return null;
        }

        return new CloudPcStatusDetail(code, message, additionalInfo);
    }

    public async Task<IReadOnlyList<CloudPcDiskSpace>> GetCloudPcDiskSpacesAsync(IReadOnlyList<CloudPcSummary>? cloudPcs = null)
    {
        cloudPcs ??= await GetCloudPcsAsync();

        var results = await ConcurrencyHelper.MapWithConcurrencyAsync(cloudPcs, maxConcurrency: 5, async cloudPc =>
        {
            ManagedDeviceDiskInfo? managedDevice = null;
            string? error = null;
            if (!string.IsNullOrWhiteSpace(cloudPc.ManagedDeviceId))
            {
                var escapedManagedDeviceId = Uri.EscapeDataString(cloudPc.ManagedDeviceId);
                try
                {
                    managedDevice = await GetAsync<ManagedDeviceDiskInfo>(
                        $"deviceManagement/managedDevices/{escapedManagedDeviceId}?$select=id,deviceName,totalStorageSpaceInBytes,freeStorageSpaceInBytes,lastSyncDateTime");
                }
                catch (HttpRequestException ex)
                {
                    error = $"Managed device disk data unavailable: {ex.Message}";
                }
            }

            var totalGb = ToGb(managedDevice?.TotalStorageSpaceInBytes);
            var freeGb = ToGb(managedDevice?.FreeStorageSpaceInBytes);
            double? usedGb = totalGb is not null && freeGb is not null
                ? Math.Round(totalGb.Value - freeGb.Value, 2)
                : null;
            double? percentFree = totalGb is > 0 && freeGb is not null
                ? Math.Round((freeGb.Value / totalGb.Value) * 100, 1)
                : null;

            return new CloudPcDiskSpace
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
                LastSyncDateTime = managedDevice?.LastSyncDateTime,
                Error = error
            };
        });

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
        // provisioningType on cloudPcServicePlan is an evolvable enum -- sharedByUser,
        // sharedByEntraGroup, and reserve only come back as their real string values with the
        // Prefer: include-unknown-enum-members header; without it Graph collapses them all to
        // "unknownFutureValue". Omitting this made the create-policy wizard's license pre-check
        // wrongly report "no purchased Flex Shared license" for tenants that actually have one,
        // since every sharedByEntraGroup plan looked identical to an unrecognized enum value.
        var plans = await GetPagedAsync<CloudPcServicePlan>(
            "deviceManagement/virtualEndpoint/servicePlans",
            includeUnknownEnumMembers: true);

        return plans
            .OrderBy(plan => plan.Type, StringComparer.OrdinalIgnoreCase)
            .ThenBy(plan => plan.VCpuCount)
            .ThenBy(plan => plan.RamGb)
            .ThenBy(plan => plan.StorageGb)
            .ThenBy(plan => plan.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    /// <summary>
    /// Windows 365 Flex ("Frontline") shared-license capacity — used when picking a service
    /// plan/license pool and remaining Cloud PC count for a sharedByEntraGroup provisioning policy
    /// assignment. Endpoint confirmed via captured browser network traffic from the real admin
    /// portal; not (yet) in the official Graph API reference.
    /// </summary>
    public async Task<IReadOnlyList<FrontLineServicePlan>> GetFrontLineServicePlansAsync()
    {
        return await GetPagedAsync<FrontLineServicePlan>("deviceManagement/virtualEndpoint/frontLineServicePlans");
    }

    /// <summary>
    /// For a specific Frontline license pool, finds every EXISTING Windows 365 Flex Dedicated
    /// (sharedByUser) policy assignment that already draws from it, and how many actual Cloud PCs
    /// each one has provisioned so far — so the create-policy wizard can surface "paid for but
    /// unused" dedicated capacity. Each reserved license unit covers up to 3 dedicated Cloud PCs;
    /// a policy that reserved 1 unit but only provisioned 1 Cloud PC still has 2 unused slots that
    /// frontLineServicePlans' own totalCount/usedCount numbers don't show at all (those only track
    /// whole reserved units, never how fully each one is actually used).
    ///
    /// Deliberately re-fetches each candidate policy's assignments via the dedicated
    /// .../provisioningPolicies/{id}/assignments navigation endpoint rather than trusting the
    /// $expand=assignments already present on the plain policy list -- undocumented beta-only
    /// fields like servicePlanId/allotmentLicensesCount on the assignment target have been
    /// observed to not survive collection-level $expand serialization, only showing up reliably
    /// via the assign POST itself and this per-item navigation endpoint.
    /// </summary>
    public async Task<int> GetUnusedDedicatedSlotsForServicePlanAsync(string servicePlanId)
    {
        var policies = await GetProvisioningPoliciesAsync();
        var dedicatedPolicyCandidates = policies
            .Where(policy => string.Equals(policy.ProvisioningType, "sharedByUser", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (dedicatedPolicyCandidates.Length == 0)
        {
            return 0;
        }

        var perPolicyUnused = await ConcurrencyHelper.MapWithConcurrencyAsync(dedicatedPolicyCandidates, maxConcurrency: 5, async policy =>
        {
            try
            {
                var assignmentsPage = await GetAsync<GraphPage<JsonElement>>(
                    $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policy.Id)}/assignments");

                var reservedUnits = 0;
                var matched = false;
                foreach (var assignment in assignmentsPage?.Value ?? [])
                {
                    if (!assignment.TryGetProperty("target", out var target))
                    {
                        continue;
                    }

                    var targetServicePlanId = GetString(target, "servicePlanId");
                    if (!string.Equals(targetServicePlanId, servicePlanId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    matched = true;
                    reservedUnits += GetInt(target, "allotmentLicensesCount") ?? 0;
                }

                if (!matched)
                {
                    return 0;
                }

                var cloudPcs = await GetCloudPcsByProvisioningPolicyAsync(policy.Id);
                return Math.Max(0, reservedUnits * 3 - cloudPcs.Count);
            }
            catch
            {
                // Best-effort only — if we can't verify a policy's actual Cloud PC count, don't
                // let one failure block showing capacity info for the rest.
                return 0;
            }
        });

        return perPolicyUnused.Sum();
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
        using var response = await SendWithRetryAsync(() => new HttpRequestMessage(HttpMethod.Delete, $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}"));
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
        }
    }

    public async Task UnassignProvisioningPolicyAsync(string policyId)
    {
        await PostJsonAsync($"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}/assign", new
        {
            assignments = Array.Empty<object>()
        });
    }

    public async Task ApplyProvisioningPolicyAsync(string policyId, int reservePercentage, bool isForceUserLogoffEnabled = false)
    {
        var body = new Dictionary<string, object>
        {
            ["reservePercentage"] = reservePercentage,
            ["isForceUserLogoffEnabled"] = isForceUserLogoffEnabled
        };

        await PostJsonAsync($"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}/apply", body);
    }

    public async Task<CloudPcPolicyApplyActionResult?> GetProvisioningPolicyApplyActionResultAsync(string policyId)
    {
        return await GetAsync<CloudPcPolicyApplyActionResult>(
            $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}/retrievePolicyApplyActionResult");
    }

    /// <summary>
    /// Resolves the assignment ID and user-settings-persistence configuration ID needed to query
    /// "user experience sync" storage usage/profiles for a shared provisioning policy. Returns
    /// null if the policy has no assignment with user settings persistence configured.
    ///
    /// Note: <c>userSettingsPersistenceDetail</c> is a navigation property on
    /// <c>cloudPcProvisioningPolicyAssignment</c> that does not support <c>$expand</c> (Graph
    /// returns "The query specified in the URI is not valid" if you try) — it must be fetched with
    /// a separate GET per assignment.
    /// </summary>
    public async Task<ProvisioningPolicyUserSettingsPersistenceContext?> GetUserSettingsPersistenceContextAsync(string policyId)
    {
        var select = Uri.EscapeDataString("id,userSettingsPersistenceConfiguration");
        var policyElement = await GetAsync<JsonElement>(
            $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}?$select={select}&$expand=assignments");

        var enabled = policyElement.TryGetProperty("userSettingsPersistenceConfiguration", out var configuration) &&
            configuration.ValueKind == JsonValueKind.Object &&
            GetBool(configuration, "userSettingsPersistenceEnabled") == true;

        double? storageSizeGb = configuration.ValueKind == JsonValueKind.Object
            ? GetString(configuration, "userSettingsPersistenceStorageSizeCategory") switch
            {
                "fourGB" => 4,
                "eightGB" => 8,
                "sixteenGB" => 16,
                "thirtyTwoGB" => 32,
                "sixtyFourGB" => 64,
                _ => null
            }
            : null;

        if (!policyElement.TryGetProperty("assignments", out var assignments) || assignments.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var assignment in assignments.EnumerateArray())
        {
            var assignmentId = GetString(assignment, "id");
            if (string.IsNullOrWhiteSpace(assignmentId))
            {
                continue;
            }

            JsonElement detail;
            try
            {
                detail = await GetAsync<JsonElement>(
                    $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(policyId)}/assignments/{Uri.EscapeDataString(assignmentId)}/userSettingsPersistenceDetail");
            }
            catch (Exception)
            {
                // Not every assignment has user settings persistence configured — Graph returns
                // a 404/400 for those; treat as "not configured" for this assignment and continue.
                continue;
            }

            var configurationId = GetString(detail, "id");
            if (string.IsNullOrWhiteSpace(configurationId))
            {
                continue;
            }

            return new ProvisioningPolicyUserSettingsPersistenceContext(policyId, assignmentId, configurationId, storageSizeGb, enabled);
        }

        return null;
    }

    public async Task<CloudPcUserSettingsPersistenceUsageResult?> GetUserSettingsPersistenceUsageAsync(ProvisioningPolicyUserSettingsPersistenceContext context)
    {
        var configurationId = context.ConfigurationId.Replace("'", "''", StringComparison.Ordinal);
        return await GetAsync<CloudPcUserSettingsPersistenceUsageResult>(
            $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(context.PolicyId)}/assignments/{Uri.EscapeDataString(context.AssignmentId)}/userSettingsPersistenceDetail/retrieveUserSettingsPersistenceProfileUsage(configurationId='{configurationId}')");
    }

    public async Task<IReadOnlyList<GraphTableRow>> GetUserSettingsPersistenceProfilesAsync(ProvisioningPolicyUserSettingsPersistenceContext context)
    {
        var configurationId = context.ConfigurationId.Replace("'", "''", StringComparison.Ordinal);
        return await GetJsonRowsAsync(
            $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(context.PolicyId)}/assignments/{Uri.EscapeDataString(context.AssignmentId)}/userSettingsPersistenceDetail/retrieveUserSettingsPersistenceProfiles(configurationId='{configurationId}')?$top=100",
            "userPrincipalName", "profileSizeInGB", "status", "lastProfileAttachedDateTime");
    }

    /// <summary>
    /// Deletes ("cleans up") a user's user settings persistence (UES) cloud profile/disk. Matches
    /// the Intune portal's delete action for a profile row on a shared policy's user experience
    /// sync screen. The Graph function accepts a batch of profile IDs but the CLI always sends one.
    /// </summary>
    public async Task BatchCleanupUserSettingsPersistenceProfileAsync(ProvisioningPolicyUserSettingsPersistenceContext context, string profileId)
    {
        await PostJsonAsync(
            $"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(context.PolicyId)}/assignments/{Uri.EscapeDataString(context.AssignmentId)}/userSettingsPersistenceDetail/batchCleanupUserSettingsPersistenceProfile",
            new
            {
                cloudProfileIds = new[] { profileId },
                configurationId = context.ConfigurationId
            });
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

    public async Task<string> CreateProvisioningPolicyAsync(
        string displayName,
        string? description,
        string provisioningType,
        string imageId,
        string imageDisplayName,
        string imageType,
        string domainJoinType,
        string? regionName,
        string? cloudPcNamingTemplate,
        bool enableSingleSignOn,
        bool localAdminEnabled,
        string? assignGroupId,
        string? userExperienceType = null,
        bool? userSettingsPersistenceEnabled = null,
        string? userSettingsPersistenceStorageSizeCategory = null,
        string? frontLineServicePlanId = null,
        string? allotmentDisplayName = null,
        int? allotmentLicensesCount = null)
    {
        var domainJoinConfiguration = new Dictionary<string, object?>
        {
            ["domainJoinType"] = domainJoinType
        };

        if (!string.IsNullOrWhiteSpace(regionName))
        {
            domainJoinConfiguration["regionName"] = regionName;
        }

        var body = new Dictionary<string, object?>
        {
            ["@odata.type"] = "#microsoft.graph.cloudPcProvisioningPolicy",
            ["displayName"] = displayName,
            ["description"] = description ?? string.Empty,
            ["provisioningType"] = provisioningType,
            // Explicitly sending userExperienceType (rather than only when it's the non-default
            // "cloudApp") matches a captured, working Windows 365 admin portal payload for a real
            // sharedByEntraGroup policy, which always includes it even for the default "cloudPc".
            ["userExperienceType"] = userExperienceType ?? "cloudPc",
            ["imageId"] = imageId,
            ["imageDisplayName"] = imageDisplayName,
            ["imageType"] = imageType,
            ["domainJoinConfigurations"] = new[] { domainJoinConfiguration },
            ["enableSingleSignOn"] = enableSingleSignOn,
            ["localAdminEnabled"] = localAdminEnabled
        };

        // Confirmed via a captured working portal payload: cloudPcNamingTemplate IS sent (as a
        // real %RAND%-based value, not null) even for sharedByEntraGroup policies -- only the
        // %USERNAME% macro is unsupported there, not the property itself.
        if (!string.IsNullOrWhiteSpace(cloudPcNamingTemplate))
        {
            body["cloudPcNamingTemplate"] = cloudPcNamingTemplate;
        }

        // userSettingsPersistenceConfiguration is documented as "only available for
        // sharedByEntraGroup" with a default of null -- but a captured working create payload for
        // a real sharedByUser (Flex Dedicated) policy sends it too (disabled, but present), so
        // it's sent for both Flex license modes, not just Shared. Plus the same-named top-level
        // property, confirmed present in both captured working create requests alongside the
        // nested object, is needed for the feature toggle in the wizard to actually take effect.
        var isFlexLicense = string.Equals(provisioningType, "sharedByEntraGroup", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(provisioningType, "sharedByUser", StringComparison.OrdinalIgnoreCase);
        if (isFlexLicense)
        {
            body["userSettingsPersistenceEnabled"] = userSettingsPersistenceEnabled ?? false;
            body["userSettingsPersistenceConfiguration"] = new Dictionary<string, object?>
            {
                ["userSettingsPersistenceEnabled"] = userSettingsPersistenceEnabled ?? false,
                ["userSettingsPersistenceStorageSizeCategory"] = userSettingsPersistenceStorageSizeCategory ?? "fourGB"
            };
        }

        var created = await PostJsonForElementAsync("deviceManagement/virtualEndpoint/provisioningPolicies", body);
        var createdId = GetString(created, "id");
        if (string.IsNullOrWhiteSpace(createdId))
        {
            throw new InvalidOperationException("Graph did not return the new provisioning policy id.");
        }

        if (!string.IsNullOrWhiteSpace(assignGroupId))
        {
            var assignmentTarget = new Dictionary<string, object?>
            {
                ["groupId"] = assignGroupId
            };

            if (isFlexLicense)
            {
                // Both Windows 365 Flex modes (Dedicated=sharedByUser, Shared=sharedByEntraGroup)
                // draw from a specific Frontline license pool (servicePlanId) and reserve a fixed
                // number of license units from it (allotmentLicensesCount), with a friendly
                // allotmentDisplayName shown in the end user's Windows app -- this whole shape is
                // undocumented in the official Graph API reference, confirmed only via captured
                // browser network traffic from the real admin portal for BOTH provisioning types.
                // Deliberately no "@odata.type" here either, matching that capture exactly (unlike
                // the plain Enterprise "dedicated" assignment shape below, which does send
                // @odata.type per the documented cloudPcManagementGroupAssignmentTarget).
                assignmentTarget["servicePlanId"] = frontLineServicePlanId;
                assignmentTarget["allotmentDisplayName"] = allotmentDisplayName ?? string.Empty;
                assignmentTarget["allotmentLicensesCount"] = allotmentLicensesCount ?? 1;
            }
            else
            {
                assignmentTarget["@odata.type"] = "#microsoft.graph.cloudPcManagementGroupAssignmentTarget";
            }

            var assignment = new Dictionary<string, object?>
            {
                ["target"] = assignmentTarget
            };

            // Confirmed via capture: both Flex modes' assignments send an empty "id" string, not
            // the "{policyId}_{groupId}" composite id used for plain Enterprise "dedicated"
            // assignments. Every other provisioning type omits "id" entirely, matching
            // pre-existing behavior.
            if (isFlexLicense)
            {
                assignment["id"] = string.Empty;
            }
            else if (string.Equals(provisioningType, "dedicated", StringComparison.OrdinalIgnoreCase))
            {
                assignment["id"] = $"{createdId}_{assignGroupId}";
            }

            await PostJsonAsync($"deviceManagement/virtualEndpoint/provisioningPolicies/{Uri.EscapeDataString(createdId)}/assign", new
            {
                assignments = new[] { assignment }
            });
        }

        return createdId;
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

    public async Task<IReadOnlyList<GraphTableRow>> GetSignInStatusRowsAsync()
    {
        var cloudPcs = await GetCloudPcsAsync();
        var rows = await ConcurrencyHelper.MapWithConcurrencyAsync(cloudPcs, maxConcurrency: 5, GetSignInStatusRowAsync);
        return rows.OrderBy(row => row.Title, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public async Task<GraphTableRow> GetSignInStatusRowAsync(CloudPcSummary cloudPc)
    {
        try
        {
            var cloudPcId = cloudPc.Id.Replace("'", "''", StringComparison.Ordinal);
            var report = await GetAsync<JsonElement>(
                $"deviceManagement/virtualEndpoint/reports/getRealTimeRemoteConnectionStatus(cloudPcId='{cloudPcId}')");
            var reportRows = ParseReportRows(report, "ManagedDeviceName", "CloudPcId", "SignInStatus", "LastActiveTime");
            var row = reportRows.FirstOrDefault();
            if (row is not null)
            {
                var fields = new Dictionary<string, string>(row.Fields, StringComparer.OrdinalIgnoreCase)
                {
                    ["Cloud PC"] = cloudPc.Name,
                    ["Cloud PC ID"] = cloudPc.Id,
                    ["User"] = cloudPc.UserPrincipalName ?? "-",
                    ["Service plan"] = cloudPc.ServicePlanName ?? "-",
                    ["Provisioning type"] = cloudPc.ProvisioningType ?? "-"
                };
                return ToTableRow(fields, "SignInStatus", "DaysSinceLastSignIn", "LastActiveTime");
            }

            return CreateUnavailableSignInRow(cloudPc, "Real-time status returned no rows");
        }
        catch (HttpRequestException ex)
        {
            return CreateUnavailableSignInRow(cloudPc, ex.Message);
        }
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

    public async Task<IReadOnlyList<GraphTableRow>> GetConnectionHistoryReportAsync(CloudPcSummary cloudPc, int top, int skip = 0)
    {
        var cloudPcId = cloudPc.Id.Replace("'", "''", StringComparison.Ordinal);
        var filter =
            "(TimeRange eq 'Last 7 days') and (PolicyNameParam eq '') and (RegionParam eq '') and " +
            "(UserSettingNameParam eq '') and (ServicePlanTypeParam eq '') and (ServicePlanNameParam eq '') and " +
            "(OSBuildVersionParam eq '') and (AADJoinTypeParam eq '') and (ImageNameParam eq '') and " +
            "(GatewayRegionParam eq '') and (ClientOSParam eq '') and (ClientTypeParam eq '') and " +
            "(ClientFriendlyNameParam eq '') and (TransportTypeParam eq '') and (CloudPCEndpointCountryRegionParam eq '') and " +
            "(CloudPCEndpointStateParam eq '') and (CloudPCEndpointCityParam eq '') and (CloudPCStatusParam eq '') and " +
            "(OSVersionParam eq '') and (ClientVersionParam eq '') and " +
            $"(CloudPCIdParam eq '{cloudPcId}') and " +
            "(UPNParam eq '') and (MMRVersionParam eq '') and (TeamsAppV2VersionParam eq '')";

        var body = new Dictionary<string, object>
        {
            ["reportName"] = "troubleshootConnectionConfigurationOfViewDataTableV1Report",
            ["top"] = top,
            ["skip"] = skip,
            ["search"] = "",
            ["filter"] = filter,
            ["select"] = new[]
            {
                "ActivityId", "SessionBeginTime", "SessionEndTime", "UPN", "UserId", "ManagedDeviceName",
                "CloudPCId", "CloudPCHostGeography", "Region", "SessionHostAgentVersion", "SessionHostSxSStackVersion",
                "GatewayRegion", "PolicyName", "UserSettingName", "ServicePlanType", "ServicePlanName", "AADJoinType",
                "ImageName", "TransportType", "PlatformName", "ClientOS", "ClientType", "ClientVersion",
                "SessionHostIPAddress", "CallerIPAddress", "CloudPCEndpointCountry", "CloudPCEndpointState",
                "CloudPCEndpointCity", "TeamsAppV2Version", "MMRVersion"
            },
            ["orderBy"] = new[] { "SessionBeginTime desc" }
        };

        var json = await PostJsonForStringAsync("deviceManagement/virtualEndpoint/reports/retrieveCloudPcTroubleshootReports", body);
        using var document = JsonDocument.Parse(json);
        var rows = ParseReportRows(document.RootElement, "UPN", "SessionBeginTime", "SessionEndTime");

        return rows
            .Select(row =>
            {
                var fields = new Dictionary<string, string>(row.Fields, StringComparer.OrdinalIgnoreCase)
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
        var rows = await ConcurrencyHelper.MapWithConcurrencyAsync(cloudPcs, maxConcurrency: 5, async cloudPc =>
        {
            if (string.IsNullOrWhiteSpace(cloudPc.UserPrincipalName))
            {
                return new GraphTableRow(
                    cloudPc.Name,
                    "Status: Skipped | Reason: No user principal name",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Cloud PC ID"] = cloudPc.Id,
                        ["Status"] = "Skipped",
                        ["Reason"] = "No user principal name"
                    });
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
                    return new GraphTableRow(
                        cloudPc.Name,
                        JoinSummary($"Status: {status}", $"Switch compatible: {switchCompatible}"),
                        fields);
                }

                // Graph returned a non-object body (e.g. empty) — treat the same as "no launch
                // detail available" rather than silently dropping the Cloud PC from the list.
                return new GraphTableRow(
                    cloudPc.Name,
                    "Status: Unavailable | Reason: No launch detail returned",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Cloud PC ID"] = cloudPc.Id,
                        ["User"] = cloudPc.UserPrincipalName,
                        ["Status"] = "Unavailable",
                        ["Reason"] = "No launch detail returned"
                    });
            }
            catch (HttpRequestException ex)
            {
                var reason = FormatLaunchDetailError(ex);
                return new GraphTableRow(
                    cloudPc.Name,
                    $"Status: Unavailable | Reason: {reason}",
                    new Dictionary<string, string>
                    {
                        ["Cloud PC"] = cloudPc.Name,
                        ["Cloud PC ID"] = cloudPc.Id,
                        ["User"] = cloudPc.UserPrincipalName,
                        ["Status"] = "Unavailable",
                        ["Reason"] = reason
                    });
            }
        });

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

        return ParseReportRows(document.RootElement, "CloudPcName", "ManagedDeviceName", "DisplayName", "SignInStatus", "Status", "Timestamp", "LastActiveTime");
    }

    private static IReadOnlyList<GraphTableRow> ParseReportRows(JsonElement report, params string[] summaryFields)
    {
        if (!report.TryGetProperty("Schema", out var schema) ||
            !report.TryGetProperty("Values", out var values) ||
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

            rows.Add(ToTableRow(fields, summaryFields));
        }

        return rows;
    }

    private static GraphTableRow CreateUnavailableSignInRow(CloudPcSummary cloudPc, string reason)
    {
        return new GraphTableRow(
            cloudPc.Name,
            JoinSummary("Status unavailable", reason),
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Cloud PC"] = cloudPc.Name,
                ["Cloud PC ID"] = cloudPc.Id,
                ["ManagedDeviceName"] = cloudPc.ManagedDeviceName ?? "-",
                ["User"] = cloudPc.UserPrincipalName ?? "-",
                ["Service plan"] = cloudPc.ServicePlanName ?? "-",
                ["SignInStatus"] = "Unavailable",
                ["DaysSinceLastSignIn"] = "-",
                ["LastActiveTime"] = "-",
                ["Reason"] = reason
            });
    }

    private async Task<List<T>> GetPagedAsync<T>(string relativeUri, bool includeConsistencyLevel = false, bool includeUnknownEnumMembers = false)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        var output = new List<T>();
        var next = relativeUri;

        while (!string.IsNullOrWhiteSpace(next))
        {
            var currentUri = next;
            using var response = await SendWithRetryAsync(() =>
            {
                var request = new HttpRequestMessage(HttpMethod.Get, currentUri);
                if (includeConsistencyLevel)
                {
                    request.Headers.Add("ConsistencyLevel", "eventual");
                }

                if (includeUnknownEnumMembers)
                {
                    request.Headers.Add("Prefer", "include-unknown-enum-members");
                }

                return request;
            });

            if (!response.IsSuccessStatusCode)
            {
                var errorBody = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
            }

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
        using var response = await SendWithRetryAsync(() => new HttpRequestMessage(HttpMethod.Get, relativeUri));
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
        }

        var content = await response.Content.ReadAsStringAsync();
        if (string.IsNullOrWhiteSpace(content))
        {
            // Some Graph GET endpoints (e.g. retrievePolicyApplyActionResult before any apply
            // operation has ever run) return a 200/204 with no body at all.
            return default;
        }

        return JsonSerializer.Deserialize<T>(content, JsonOptions);
    }

    /// <summary>
    /// For endpoints that return a plain non-JSON body, like Graph's /$count suffix (returns just
    /// a bare number as text/plain, not JSON) -- $count also requires ConsistencyLevel: eventual
    /// per Graph's own requirements for count queries.
    /// </summary>
    private async Task<string> GetRawStringAsync(string relativeUri, bool includeConsistencyLevel = false)
    {
        using var response = await SendWithRetryAsync(() =>
        {
            var request = new HttpRequestMessage(HttpMethod.Get, relativeUri);
            if (includeConsistencyLevel)
            {
                request.Headers.Add("ConsistencyLevel", "eventual");
            }

            return request;
        });

        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
        }

        return await response.Content.ReadAsStringAsync();
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

        using var response = await SendWithRetryAsync(() => new HttpRequestMessage(HttpMethod.Post, relativeUri)
        {
            Content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), System.Text.Encoding.UTF8, "application/json")
        });
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException($"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}");
        }
    }

    private async Task<string> PostJsonForStringAsync(string relativeUri, object body)
    {
        var requestJson = JsonSerializer.Serialize(body, JsonOptions);
        using var response = await SendWithRetryAsync(() => new HttpRequestMessage(HttpMethod.Post, relativeUri)
        {
            Content = new StringContent(requestJson, System.Text.Encoding.UTF8, "application/json"),
            Headers = { { "Prefer", "include-unknown-enum-members" } }
        });
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new HttpRequestException(
                $"{(int)response.StatusCode} {response.ReasonPhrase}: {ExtractGraphErrorMessage(errorBody)}\nRequest body: {requestJson}");
        }

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

    /// <summary>
    /// Sends a Graph request with retry/backoff for throttling (429) and transient upstream
    /// errors (502/503/504) — Graph throttles aggressively on list-heavy tenants, and without
    /// this a single throttled call surfaces as a raw error instead of quietly retrying.
    ///
    /// Takes a factory rather than a single HttpRequestMessage because request messages (and any
    /// StringContent body) can only be sent once — retrying requires building a fresh message
    /// each attempt. Honors the Retry-After header when Graph provides one (mandatory on 429,
    /// often present on 503), falling back to exponential backoff with jitter otherwise. The
    /// final attempt's response (success or failure) is returned to the caller un-disposed for
    /// its own `using`; only the intermediate retried-away responses are disposed here.
    /// </summary>
    private async Task<HttpResponseMessage> SendWithRetryAsync(Func<HttpRequestMessage> requestFactory)
    {
        if (_accessTokenProvider is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        const int maxAttempts = 4;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            using var request = requestFactory();
            await AuthorizeAsync(request);
            var response = await _httpClient.SendAsync(request);

            if (attempt >= maxAttempts || !IsRetryableStatus(response.StatusCode))
            {
                return response;
            }

            var delay = GetRetryDelay(response, attempt);
            response.Dispose();
            await Task.Delay(delay);
        }

        // Unreachable — the loop always returns or retries within maxAttempts.
        throw new InvalidOperationException("Retry loop exited without a response.");
    }

    private static bool IsRetryableStatus(HttpStatusCode statusCode)
    {
        return statusCode is HttpStatusCode.TooManyRequests // 429 — throttled
            or HttpStatusCode.ServiceUnavailable // 503
            or HttpStatusCode.BadGateway // 502
            or HttpStatusCode.GatewayTimeout; // 504
    }

    /// <summary>
    /// Honors Graph's Retry-After header (either a delta-seconds or an absolute date) when
    /// present, capped at 30s so an interactive CLI session never hangs unreasonably long
    /// waiting on a single retry. Falls back to exponential backoff with jitter (~1s, 2s, 4s
    /// for attempts 1/2/3) when no Retry-After is given, which happens on some 502/503/504s that
    /// aren't a formal throttling response.
    /// </summary>
    private static TimeSpan GetRetryDelay(HttpResponseMessage response, int attempt)
    {
        var cap = TimeSpan.FromSeconds(30);
        if (response.Headers.RetryAfter is { } retryAfter)
        {
            if (retryAfter.Delta is { } delta && delta > TimeSpan.Zero)
            {
                return delta > cap ? cap : delta;
            }

            if (retryAfter.Date is { } date)
            {
                var wait = date - DateTimeOffset.UtcNow;
                if (wait > TimeSpan.Zero)
                {
                    return wait > cap ? cap : wait;
                }
            }
        }

        var baseDelayMs = Math.Pow(2, attempt - 1) * 1000;
        var jitterMs = Random.Shared.Next(0, 250);
        return TimeSpan.FromMilliseconds(baseDelayMs + jitterMs);
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
        var resolved = await ConcurrencyHelper.MapWithConcurrencyAsync(groupIds, maxConcurrency: 5, async groupId =>
        {
            try
            {
                var group = await GetAsync<JsonElement>($"groups/{Uri.EscapeDataString(groupId)}?$select=id,displayName");
                var name = group.ValueKind == JsonValueKind.Object ? GetString(group, "displayName") : null;
                return (groupId, name: string.IsNullOrWhiteSpace(name) ? groupId : name);
            }
            catch (HttpRequestException)
            {
                return (groupId, name: groupId);
            }
        });

        return resolved.ToDictionary(pair => pair.groupId, pair => pair.name, StringComparer.OrdinalIgnoreCase);
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

    private static string ExtractGraphErrorMessage(string errorBody)
    {
        if (string.IsNullOrWhiteSpace(errorBody))
        {
            return "No additional error details were returned.";
        }

        try
        {
            using var document = JsonDocument.Parse(errorBody);
            if (document.RootElement.TryGetProperty("error", out var error))
            {
                var code = error.TryGetProperty("code", out var codeElement) ? codeElement.GetString() : null;
                var message = error.TryGetProperty("message", out var messageElement) ? messageElement.GetString() : null;
                if (!string.IsNullOrWhiteSpace(code) || !string.IsNullOrWhiteSpace(message))
                {
                    return string.Join(": ", new[] { code, message }.Where(value => !string.IsNullOrWhiteSpace(value)));
                }
            }
        }
        catch (JsonException)
        {
            // Fall through to returning the raw body below.
        }

        return errorBody;
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
