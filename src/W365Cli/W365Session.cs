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

    public string? TenantName { get; private set; }

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

        try
        {
            _credential = new InteractiveBrowserCredential(options);
            var token = await _credential.GetTokenAsync(new TokenRequestContext(_scopes), CancellationToken.None);

            Graph = new W365GraphClient(_credential, _scopes);
            IsConnected = !string.IsNullOrWhiteSpace(token.Token);
            if (IsConnected)
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

            AnsiConsole.MarkupLine(IsConnected ? "[green]Connected.[/]" : "[red]Connection failed.[/]");
        }
        catch (AuthenticationFailedException ex)
        {
            IsConnected = false;
            Graph = W365GraphClient.NotConnected;
            _credential = null;

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

    public void Disconnect()
    {
        _credential = null;
        Graph = W365GraphClient.NotConnected;
        IsConnected = false;
        TenantId = null;
        TenantName = null;
    }
}
