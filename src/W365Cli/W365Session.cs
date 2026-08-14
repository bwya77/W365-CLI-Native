using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Spectre.Console;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace W365Cli;

internal sealed class W365Session
{
    private const string DefaultClientId = "9d497858-c200-402c-a363-279a5800d730";

    private readonly string[] _scopes =
    [
        "https://graph.microsoft.com/.default"
    ];

    /// <summary>
    /// Every delegated Microsoft Graph permission this CLI relies on somewhere in its feature set.
    /// This is the single source of truth for "what should be granted on the app registration" —
    /// when you add a Graph call that needs a permission not already covered here, add it to this
    /// list AND to the app registration's API permissions in Entra (Add a permission → grant admin
    /// consent). The list drives <see cref="MissingRequiredScopes"/>, which proactively warns users
    /// right after connecting instead of letting them hit a confusing 403 mid-action.
    /// </summary>
    private static readonly string[] RequiredScopes =
    [
        "CloudPC.ReadWrite.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementManagedDevices.PrivilegedOperations.All",
        "Group.Read.All",
        "GroupMember.ReadWrite.All",
        "User.Read.All",
        "Organization.Read.All"
    ];

    private IPublicClientApplication? _application;
    private AuthenticationResult? _currentAuthentication;
    private string _clientId = DefaultClientId;

    public bool IsConnected { get; private set; }

    public string? TenantId { get; private set; }

    public string? TenantName { get; private set; }

    /// <summary>
    /// The signed-in account's UPN (e.g. "alice@contoso.com"), from MSAL's cached account —
    /// surfaced in the header's "Signed in as" line. Null when not connected.
    /// </summary>
    public string? SignedInUserUpn => _currentAuthentication?.Account?.Username;

    /// <summary>
    /// Any permissions from <see cref="RequiredScopes"/> that are NOT present in the granted token
    /// scopes after connecting — meaning they likely haven't been added to the app registration
    /// and/or admin-consented in this tenant yet. Empty when everything the app needs is granted.
    /// Used to proactively surface a "grant permission" prompt on first connect instead of letting
    /// the user hit a confusing 403 later, mid-action.
    /// </summary>
    public IReadOnlyList<string> MissingRequiredScopes { get; private set; } = [];

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
            UpdateMissingPermissionFlag();
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
            _currentAuthentication = IsRunningInWsl()
                ? await AcquireTokenViaDeviceCodeAsync()
                : await _application
                    .AcquireTokenInteractive(_scopes)
                    .WithPrompt(Prompt.SelectAccount)
                    .ExecuteAsync();

            ConfigureConnectedGraph();
            UpdateMissingPermissionFlag();
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
        catch (Exception ex)
        {
            IsConnected = false;
            Graph = W365GraphClient.NotConnected;
            _currentAuthentication = null;

            AnsiConsole.MarkupLine("[red]Authentication failed.[/]");
            AnsiConsole.MarkupLine(Markup.Escape(ex.Message));
        }
    }

    /// <summary>
    /// Forces a fresh interactive sign-in with the Microsoft consent screen shown again, even if a
    /// cached token already exists. Use this to let the user (or a tenant admin) grant/consent to
    /// permissions that weren't accepted the first time around.
    /// </summary>
    public async Task<bool> ReconsentAsync()
    {
        try
        {
            _application ??= await CreateApplicationAsync();
            _currentAuthentication = IsRunningInWsl()
                ? await AcquireTokenViaDeviceCodeAsync()
                : await _application
                    .AcquireTokenInteractive(_scopes)
                    .WithPrompt(Prompt.Consent)
                    .ExecuteAsync();

            ConfigureConnectedGraph();
            UpdateMissingPermissionFlag();
            return IsConnected;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Sign-in path for WSL. The normal interactive flow relies on Microsoft redirecting the
    /// browser back to a loopback HTTP listener the app spins up locally -- but that listener runs
    /// inside the WSL2 Linux VM while the browser we launch (via explorer.exe interop) is the
    /// user's real Windows browser. WSL2's automatic localhost port-forwarding is supposed to
    /// bridge that gap, but in practice it's unreliable (VPN adapters, firewall rules, and
    /// IPv4/IPv6 loopback mismatches all break it) -- which is exactly what happened here: sign-in
    /// completed fine in the browser, but the app never saw the redirect and hung forever. Device
    /// code flow sidesteps the whole problem: no listener, no redirect back to WSL at all -- just
    /// a short code the user enters on a Microsoft page, polled from here until they finish.
    /// </summary>
    private async Task<AuthenticationResult> AcquireTokenViaDeviceCodeAsync()
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Not connected to Microsoft Graph.");
        }

        return await _application
            .AcquireTokenWithDeviceCode(_scopes, deviceCodeResult =>
            {
                AnsiConsole.MarkupLine("[yellow]WSL detected — using device sign-in instead of a browser redirect (the WSL↔Windows loopback isn't reliable for this).[/]");
                AnsiConsole.MarkupLine(Markup.Escape(deviceCodeResult.Message));

                // deviceCodeResult.VerificationUrl is a short, plain microsoft.com/devicelogin
                // link with no query-string encoding to worry about, so the same explorer.exe
                // launch used elsewhere for WSL is safe here -- best-effort only; if it fails the
                // user still has the URL printed above to open by hand.
                try
                {
                    OpenUrlOnWindowsSide(deviceCodeResult.VerificationUrl);
                }
                catch
                {
                    // Best-effort only — the printed URL/code above still let the user sign in manually.
                }

                return Task.CompletedTask;
            })
            .ExecuteAsync();
    }


    /// <summary>
    /// Best-effort URL that lets a tenant administrator grant admin consent for this app
    /// registration's requested permissions without having to hunt through the Entra portal.
    /// Includes redirect_uri (matching the app's registered "http://localhost" redirect) — without
    /// it, Microsoft's /adminconsent endpoint can fail to redirect back after consent completes.
    /// </summary>
    public string GetAdminConsentUrl()
    {
        var tenantSegment = string.IsNullOrWhiteSpace(TenantId) ? "common" : TenantId;
        var redirectUri = Uri.EscapeDataString("http://localhost");
        return $"https://login.microsoftonline.com/{tenantSegment}/adminconsent?client_id={_clientId}&redirect_uri={redirectUri}";
    }

    /// <summary>
    /// True when running inside WSL. WSL has no xdg-open/gnome-open/kfmclient (and usually no
    /// wslview either) by default, so MSAL's built-in Linux system-browser launcher fails with
    /// "Unable to open a web page using xdg-open, gnome-open, kfmclient or wslview tools" even
    /// though a real Windows browser is one hop away via interop.
    /// </summary>
    private static bool IsRunningInWsl()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            return false;
        }

        try
        {
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WSL_DISTRO_NAME")))
            {
                return true;
            }

            const string VersionPath = "/proc/version";
            if (File.Exists(VersionPath))
            {
                var versionText = File.ReadAllText(VersionPath);
                return versionText.Contains("microsoft", StringComparison.OrdinalIgnoreCase);
            }
        }
        catch
        {
            // Best-effort detection only — fall through to "not WSL" on any failure.
        }

        return false;
    }

    /// <summary>
    /// On WSL, hands a URL off to the Windows side by launching explorer.exe directly (present on
    /// every WSL install via interop), which forwards it to the default Windows browser. Used only
    /// for the plain, unencoded device-code verification link -- deliberately not going through
    /// cmd.exe, since cmd.exe re-parses whatever raw command line it receives using its own rules
    /// (an unquoted "&" is a command separator, and "%" is an environment-variable delimiter even
    /// inside quotes), which silently mangles anything with query-string encoding.
    /// </summary>
    private static void OpenUrlOnWindowsSide(string url)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "explorer.exe",
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        startInfo.ArgumentList.Add(url);
        using var process = Process.Start(startInfo);
    }

    private void UpdateMissingPermissionFlag()
    {
        if (!IsConnected)
        {
            MissingRequiredScopes = [];
            return;
        }

        var grantedScopes = _currentAuthentication?.Scopes ?? [];

        // Granted scope strings come back like "CloudPC.ReadWrite.All" or with a resource prefix
        // ("https://graph.microsoft.com/CloudPC.ReadWrite.All") depending on the flow — compare on
        // just the trailing permission name so both shapes match correctly.
        bool IsGranted(string requiredScope) => grantedScopes.Any(granted =>
            granted.Equals(requiredScope, StringComparison.OrdinalIgnoreCase) ||
            granted.EndsWith($"/{requiredScope}", StringComparison.OrdinalIgnoreCase));

        MissingRequiredScopes = RequiredScopes.Where(scope => !IsGranted(scope)).ToArray();
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
        MissingRequiredScopes = [];
    }

    private async Task<IPublicClientApplication> CreateApplicationAsync()
    {
        var clientId = Environment.GetEnvironmentVariable("W365CLI_CLIENT_ID");
        if (string.IsNullOrWhiteSpace(clientId))
        {
            clientId = DefaultClientId;
        }

        _clientId = clientId;

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

        const string CacheFileName = "w365cli-native.msalcache";
        var storageBuilder = new StorageCreationPropertiesBuilder(CacheFileName, cacheDirectory);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            storageBuilder = storageBuilder.WithMacKeyChain("com.bwya77.w365cli", "MSALCache");
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            // Uses the system's libsecret-compatible keyring (GNOME Keyring, KWallet via
            // libsecret, etc.) when one is available. Not every Linux environment has a keyring
            // daemon running (e.g. some headless servers or minimal containers), so this is
            // verified below and falls back to a plain ACL-restricted file if it doesn't work,
            // rather than silently storing tokens unencrypted with no protection at all.
            storageBuilder = storageBuilder.WithLinuxKeyring(
                schemaName: "com.bwya77.w365cli.tokencache",
                collection: "default",
                secretLabel: "W365 CLI MSAL token cache",
                attribute1: new KeyValuePair<string, string>("Version", "1"),
                attribute2: new KeyValuePair<string, string>("Product", "W365CLI"));
        }

        var storageProperties = storageBuilder.Build();
        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            try
            {
                cacheHelper.VerifyPersistence();
            }
            catch (MsalCachePersistenceException)
            {
                // No usable keyring on this machine — fall back to an unprotected, ACL-restricted
                // file (readable only by the current user) instead of failing sign-in entirely.
                // This matches MSAL's documented plain-text fallback pattern.
                AnsiConsole.MarkupLine("[yellow]No secure keyring is available on this Linux machine — the sign-in cache will be stored in a plain, user-only-readable file instead.[/]");
                var fallbackProperties = new StorageCreationPropertiesBuilder(CacheFileName + ".plaintext", cacheDirectory)
                    .WithUnprotectedFile()
                    .Build();
                cacheHelper = await MsalCacheHelper.CreateAsync(fallbackProperties);
            }
        }

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
