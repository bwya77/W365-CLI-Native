using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Spectre.Console;

namespace W365Cli;

internal sealed class W365Session
{
    private const string DefaultClientId = "9d497858-c200-402c-a363-279a5800d730";

    private readonly string[] _scopes =
    [
        "https://graph.microsoft.com/.default"
    ];

    private IPublicClientApplication? _application;
    private AuthenticationResult? _currentAuthentication;

    public bool IsConnected { get; private set; }

    public string? TenantId { get; private set; }

    public string? TenantName { get; private set; }

    public W365GraphClient Graph { get; private set; } = W365GraphClient.NotConnected;

    public async Task TryRestoreAsync()
    {
        try
        {
            _application = await CreateApplicationAsync();
            var account = (await _application.GetAccountsAsync()).FirstOrDefault();
            if (account is null)
            {
                return;
            }

            _currentAuthentication = await _application.AcquireTokenSilent(_scopes, account).ExecuteAsync();
            ConfigureConnectedGraph();
        }
        catch
        {
            IsConnected = false;
            Graph = W365GraphClient.NotConnected;
            _currentAuthentication = null;
        }
    }

    public async Task ConnectAsync()
    {
        try
        {
            _application = await CreateApplicationAsync();
            _currentAuthentication = await _application
                .AcquireTokenInteractive(_scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync();

            ConfigureConnectedGraph();
            AnsiConsole.MarkupLine(IsConnected ? "[green]Connected.[/]" : "[red]Connection failed.[/]");
        }
        catch (MsalServiceException ex)
        {
            IsConnected = false;
            Graph = W365GraphClient.NotConnected;
            _currentAuthentication = null;

            AnsiConsole.MarkupLine("[red]Authentication failed.[/]");
            AnsiConsole.MarkupLine(Markup.Escape(ex.Message));
            if (ex.Message.Contains("AADSTS7000218", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("client_secret", StringComparison.OrdinalIgnoreCase))
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[yellow]This usually means the app registration is not configured as a public client/native app.[/]");
                AnsiConsole.MarkupLine("In Entra, enable public client flows and add a Mobile and desktop redirect URI of [grey]http://localhost[/].");
            }
        }
    }

    public async Task DisconnectAsync()
    {
        if (_application is not null)
        {
            foreach (var account in await _application.GetAccountsAsync())
            {
                await _application.RemoveAsync(account);
            }
        }

        _currentAuthentication = null;
        Graph = W365GraphClient.NotConnected;
        IsConnected = false;
        TenantId = null;
        TenantName = null;
    }

    private async Task<IPublicClientApplication> CreateApplicationAsync()
    {
        var clientId = Environment.GetEnvironmentVariable("W365CLI_CLIENT_ID");
        if (string.IsNullOrWhiteSpace(clientId))
        {
            clientId = DefaultClientId;
        }

        var tenantId = Environment.GetEnvironmentVariable("W365CLI_TENANT_ID");
        TenantId = tenantId;

        var builder = PublicClientApplicationBuilder
            .Create(clientId)
            .WithRedirectUri("http://localhost");

        builder = string.IsNullOrWhiteSpace(tenantId)
            ? builder.WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
            : builder.WithAuthority(AzureCloudInstance.AzurePublic, tenantId);

        var application = builder.Build();
        var cacheDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "W365CliNative");
        Directory.CreateDirectory(cacheDirectory);

        var storageProperties = new StorageCreationPropertiesBuilder("w365cli-native.msalcache", cacheDirectory)
            .Build();
        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
        cacheHelper.RegisterCache(application.UserTokenCache);
        return application;
    }

    private void ConfigureConnectedGraph()
    {
        IsConnected = _currentAuthentication is not null && !string.IsNullOrWhiteSpace(_currentAuthentication.AccessToken);
        Graph = new W365GraphClient(GetAccessTokenAsync);
        TenantId = _currentAuthentication?.TenantId;
        LoadTenantMetadataAsync().GetAwaiter().GetResult();
    }

    private async Task<string> GetAccessTokenAsync()
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        var account = _currentAuthentication?.Account ?? (await _application.GetAccountsAsync()).FirstOrDefault();
        if (account is null)
        {
            throw new InvalidOperationException("No cached Microsoft Graph account was found.");
        }

        _currentAuthentication = await _application.AcquireTokenSilent(_scopes, account).ExecuteAsync();
        return _currentAuthentication.AccessToken;
    }

    private async Task LoadTenantMetadataAsync()
    {
        try
        {
            var organization = await Graph.GetOrganizationAsync();
            if (organization is not null)
            {
                TenantId = organization.Id;
                TenantName = organization.DisplayName;
            }
        }
        catch
        {
            // Tenant display is helpful but not required for command execution.
        }
    }
}
