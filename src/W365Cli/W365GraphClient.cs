using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;

namespace W365Cli;

internal sealed class W365GraphClient
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly TokenCredential? _credential;
    private readonly string[] _scopes;
    private readonly HttpClient _httpClient;

    public static W365GraphClient NotConnected { get; } = new(null, []);

    public W365GraphClient(TokenCredential? credential, string[] scopes)
    {
        _credential = credential;
        _scopes = scopes;
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri("https://graph.microsoft.com/beta/")
        };
    }

    public async Task<IReadOnlyList<CloudPcSummary>> GetCloudPcsAsync()
    {
        var items = await GetPagedAsync<CloudPcSummary>(
            "deviceManagement/virtualEndpoint/cloudPCs?$select=id,displayName,managedDeviceName,status,provisioningType,userPrincipalName,servicePlanName");

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

    private async Task<List<T>> GetPagedAsync<T>(string relativeUri, bool includeConsistencyLevel = false)
    {
        if (_credential is null)
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

    private async Task AuthorizeAsync(HttpRequestMessage request)
    {
        var token = await _credential!.GetTokenAsync(new TokenRequestContext(_scopes), CancellationToken.None);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
    }
}
