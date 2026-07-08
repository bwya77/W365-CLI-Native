using System.Text.Json;
using System.Text.Json.Serialization;

namespace W365Cli;

internal sealed record GraphPage<T>(
    [property: JsonPropertyName("value")] IReadOnlyList<T> Value,
    [property: JsonPropertyName("@odata.nextLink")] string? NextLink);

internal sealed record GraphTableRow(string Title, string Summary, IReadOnlyDictionary<string, string> Fields);

internal sealed record SubscribedSku
{
    [JsonPropertyName("skuId")]
    public string? SkuId { get; init; }

    [JsonPropertyName("skuPartNumber")]
    public string? SkuPartNumber { get; init; }

    [JsonPropertyName("prepaidUnits")]
    public SubscribedSkuPrepaidUnits? PrepaidUnits { get; init; }

    [JsonPropertyName("consumedUnits")]
    public int? ConsumedUnits { get; init; }

    [JsonPropertyName("servicePlans")]
    public IReadOnlyList<SubscribedSkuServicePlan>? ServicePlans { get; init; }
}

internal sealed record SubscribedSkuPrepaidUnits
{
    [JsonPropertyName("enabled")]
    public int? Enabled { get; init; }

    [JsonPropertyName("suspended")]
    public int? Suspended { get; init; }

    [JsonPropertyName("warning")]
    public int? Warning { get; init; }
}

internal sealed record SubscribedSkuServicePlan
{
    [JsonPropertyName("servicePlanName")]
    public string? ServicePlanName { get; init; }

    [JsonPropertyName("provisioningStatus")]
    public string? ProvisioningStatus { get; init; }
}

internal sealed record GroupMemberSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; init; }

    [JsonIgnore]
    public string Name => !string.IsNullOrWhiteSpace(DisplayName)
        ? DisplayName!
        : UserPrincipalName ?? Id;
}

internal sealed record CloudPcSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }

    [JsonPropertyName("managedDeviceName")]
    public string? ManagedDeviceName { get; init; }

    [JsonPropertyName("status")]
    public string? Status { get; init; }

    [JsonPropertyName("powerState")]
    public string? PowerState { get; init; }

    [JsonPropertyName("provisioningType")]
    public string? ProvisioningType { get; init; }

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; init; }

    [JsonPropertyName("servicePlanName")]
    public string? ServicePlanName { get; init; }

    [JsonPropertyName("managedDeviceId")]
    public string? ManagedDeviceId { get; init; }

    [JsonPropertyName("provisioningPolicyId")]
    public string? ProvisioningPolicyId { get; init; }

    [JsonPropertyName("provisioningPolicyName")]
    public string? ProvisioningPolicyName { get; init; }

    [JsonIgnore]
    public string Name => !string.IsNullOrWhiteSpace(DisplayName)
        ? DisplayName!
        : ManagedDeviceName ?? Id;
}

internal sealed record ProvisioningPolicySummary
{
    public string Id { get; init; } = string.Empty;

    public string DisplayName { get; init; } = string.Empty;

    public string? Description { get; init; }

    public string? ProvisioningType { get; init; }

    public string? ImageDisplayName { get; init; }

    public string? ImageType { get; init; }

    public string? DomainJoinTypes { get; init; }

    public bool? EnableSingleSignOn { get; init; }

    public bool? LocalAdminEnabled { get; init; }

    public string? CloudPcNamingTemplate { get; init; }

    public string? CloudPcGroupDisplayName { get; init; }

    public string? ManagedBy { get; init; }

    public int? GracePeriodInHours { get; init; }

    public IReadOnlyList<string> AssignedGroupIds { get; init; } = [];

    public IReadOnlyList<string> AssignedGroupNames { get; init; } = [];

    public JsonElement Raw { get; init; }
}

internal sealed record CloudAppSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayNameValue { get; init; }

    [JsonPropertyName("appName")]
    public string? AppName { get; init; }

    [JsonPropertyName("discoveredAppName")]
    public string? DiscoveredAppName { get; init; }

    [JsonPropertyName("appStatus")]
    public string? AppStatus { get; init; }

    [JsonPropertyName("publisher")]
    public string? Publisher { get; init; }

    [JsonPropertyName("addedDateTime")]
    public DateTimeOffset? AddedDateTime { get; init; }

    [JsonPropertyName("lastPublishedDateTime")]
    public DateTimeOffset? LastPublishedDateTime { get; init; }

    [JsonIgnore]
    public string DisplayName => !string.IsNullOrWhiteSpace(DisplayNameValue)
        ? DisplayNameValue!
        : AppName ?? DiscoveredAppName ?? Id;
}

internal sealed record ManagedDeviceDiskInfo
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("deviceName")]
    public string? DeviceName { get; init; }

    [JsonPropertyName("totalStorageSpaceInBytes")]
    public long? TotalStorageSpaceInBytes { get; init; }

    [JsonPropertyName("freeStorageSpaceInBytes")]
    public long? FreeStorageSpaceInBytes { get; init; }

    [JsonPropertyName("lastSyncDateTime")]
    public DateTimeOffset? LastSyncDateTime { get; init; }
}

internal sealed record CloudPcDiskSpace
{
    public string CloudPcId { get; init; } = string.Empty;

    public string CloudPcName { get; init; } = string.Empty;

    public string? AssignedUserUpn { get; init; }

    public string? ManagedDeviceId { get; init; }

    public string? ManagedDeviceName { get; init; }

    public double? TotalStorageGb { get; init; }

    public double? FreeStorageGb { get; init; }

    public double? UsedStorageGb { get; init; }

    public double? PercentFree { get; init; }

    public DateTimeOffset? LastSyncDateTime { get; init; }
}

internal sealed record CloudPcSnapshotRaw
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("cloudPcId")]
    public string? CloudPcId { get; init; }

    [JsonPropertyName("status")]
    public string? Status { get; init; }

    [JsonPropertyName("snapshotType")]
    public string? SnapshotType { get; init; }

    [JsonPropertyName("createdDateTime")]
    public DateTimeOffset? CreatedDateTime { get; init; }

    [JsonPropertyName("expirationDateTime")]
    public DateTimeOffset? ExpirationDateTime { get; init; }

    [JsonPropertyName("lastRestoredDateTime")]
    public DateTimeOffset? LastRestoredDateTime { get; init; }

    [JsonPropertyName("healthCheckStatus")]
    public string? HealthCheckStatus { get; init; }
}

internal sealed record CloudPcSnapshot
{
    public string SnapshotId { get; init; } = string.Empty;

    public string CloudPcId { get; init; } = string.Empty;

    public string? Status { get; init; }

    public string? SnapshotType { get; init; }

    public DateTimeOffset? CreatedDateTime { get; init; }

    public DateTimeOffset? ExpirationDateTime { get; init; }

    public DateTimeOffset? LastRestoredDateTime { get; init; }

    public string? HealthCheckStatus { get; init; }
}

internal sealed record CloudPcServicePlan
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }

    [JsonPropertyName("type")]
    public string? Type { get; init; }

    [JsonPropertyName("vCpuCount")]
    public int? VCpuCount { get; init; }

    [JsonPropertyName("ramInGB")]
    public int? RamGb { get; init; }

    [JsonPropertyName("storageInGB")]
    public int? StorageGb { get; init; }

    [JsonPropertyName("userProfileInGB")]
    public int? UserProfileGb { get; init; }

    [JsonIgnore]
    public string Name => string.IsNullOrWhiteSpace(DisplayName) ? Id : DisplayName;
}

internal sealed record CloudPcRemoteActionResultRaw
{
    [JsonPropertyName("actionName")]
    public string? ActionName { get; init; }

    [JsonPropertyName("actionState")]
    public string? ActionState { get; init; }

    [JsonPropertyName("startDateTime")]
    public DateTimeOffset? StartDateTime { get; init; }

    [JsonPropertyName("lastUpdatedDateTime")]
    public DateTimeOffset? LastUpdatedDateTime { get; init; }

    [JsonPropertyName("managedDeviceId")]
    public string? ManagedDeviceId { get; init; }

    [JsonPropertyName("statusDetail")]
    public CloudPcRemoteActionStatusDetail? StatusDetail { get; init; }
}

internal sealed record CloudPcRemoteActionStatusDetail
{
    [JsonPropertyName("code")]
    public string? Code { get; init; }

    [JsonPropertyName("message")]
    public string? Message { get; init; }
}

internal sealed record CloudPcRemoteActionResult
{
    public string CloudPcId { get; init; } = string.Empty;

    public string CloudPcName { get; init; } = string.Empty;

    public string? ActionName { get; init; }

    public string? ActionState { get; init; }

    public DateTimeOffset? StartDateTime { get; init; }

    public DateTimeOffset? LastUpdatedDateTime { get; init; }

    public string? ManagedDeviceId { get; init; }

    public string? StatusCode { get; init; }

    public string? StatusMessage { get; init; }
}

internal sealed record OrganizationSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }
}
