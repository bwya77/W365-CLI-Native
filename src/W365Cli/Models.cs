using System.Text.Json.Serialization;

namespace W365Cli;

internal sealed record GraphPage<T>(
    [property: JsonPropertyName("value")] IReadOnlyList<T> Value,
    [property: JsonPropertyName("@odata.nextLink")] string? NextLink);

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

    [JsonPropertyName("provisioningType")]
    public string? ProvisioningType { get; init; }

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; init; }

    [JsonPropertyName("servicePlanName")]
    public string? ServicePlanName { get; init; }

    [JsonPropertyName("managedDeviceId")]
    public string? ManagedDeviceId { get; init; }

    [JsonIgnore]
    public string Name => !string.IsNullOrWhiteSpace(DisplayName)
        ? DisplayName!
        : ManagedDeviceName ?? Id;
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

internal sealed record OrganizationSummary
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; init; }
}
