using Azure.Core;
using Azure.Identity;
using Spectre.Console;

namespace W365Cli;

internal sealed class W365Session
{
    private const string DefaultClientId = "9d497858-c200-402c-a363-279a5800d730";

    private readonly string[] _scopes =
    [
        "https://graph.microsoft.com/.default"
    ];

    private InteractiveBrowserCredential? _credential;

    public bool IsConnected { get; private set; }

    public string? TenantId { get; private set; }

    public W365GraphClient Graph { get; private set; } = W365GraphClient.NotConnected;

    public async Task ConnectAsync()
    {
        var clientId = Environment.GetEnvironmentVariable("W365CLI_CLIENT_ID");
        if (string.IsNullOrWhiteSpace(clientId))
        {
            clientId = DefaultClientId;
        }

        TenantId = Environment.GetEnvironmentVariable("W365CLI_TENANT_ID");

        var options = new InteractiveBrowserCredentialOptions
        {
            ClientId = clientId,
            TenantId = string.IsNullOrWhiteSpace(TenantId) ? null : TenantId,
            RedirectUri = new Uri("http://localhost"),
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions
            {
                Name = "W365CliNative"
            }
        };

        _credential = new InteractiveBrowserCredential(options);
        var token = await _credential.GetTokenAsync(new TokenRequestContext(_scopes), CancellationToken.None);

        Graph = new W365GraphClient(_credential, _scopes);
        IsConnected = !string.IsNullOrWhiteSpace(token.Token);
        AnsiConsole.MarkupLine(IsConnected ? "[green]Connected.[/]" : "[red]Connection failed.[/]");
    }

    public void Disconnect()
    {
        _credential = null;
        Graph = W365GraphClient.NotConnected;
        IsConnected = false;
    }
}
