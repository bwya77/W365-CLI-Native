using Xunit;

namespace W365Cli.Tests;

public class ConnectionManagementTests
{
    private static CachedConnection MakeConnection(
        string homeAccountId,
        string username,
        string? tenantId = null,
        string? tenantName = null,
        DateTimeOffset? lastUsedUtc = null,
        bool isActive = false)
    {
        return new CachedConnection(
            homeAccountId,
            username,
            tenantId,
            tenantName,
            lastUsedUtc ?? DateTimeOffset.MinValue,
            isActive);
    }

    [Fact]
    public void OrderConnections_ActiveConnectionAlwaysFirst()
    {
        var connections = new[]
        {
            MakeConnection("1", "alice@contoso.com", lastUsedUtc: DateTimeOffset.UtcNow, isActive: false),
            MakeConnection("2", "bob@fabrikam.com", lastUsedUtc: DateTimeOffset.UtcNow.AddDays(-5), isActive: true)
        };

        var result = W365Session.OrderConnections(connections);
        Assert.Equal("bob@fabrikam.com", result[0].Username);
    }

    [Fact]
    public void OrderConnections_NonActiveSortedByMostRecentlyUsedFirst()
    {
        var connections = new[]
        {
            MakeConnection("1", "older@contoso.com", lastUsedUtc: DateTimeOffset.UtcNow.AddDays(-10)),
            MakeConnection("2", "newer@contoso.com", lastUsedUtc: DateTimeOffset.UtcNow.AddDays(-1))
        };

        var result = W365Session.OrderConnections(connections);
        Assert.Equal(["newer@contoso.com", "older@contoso.com"], result.Select(c => c.Username));
    }

    [Fact]
    public void FormatConnectionTenantLabel_PrefersTenantNameOverTenantId()
    {
        var connection = MakeConnection("1", "alice@contoso.com", tenantId: "guid-123", tenantName: "Contoso Ltd");
        Assert.Equal("Contoso Ltd", W365CliApp.FormatConnectionTenantLabel(connection));
    }

    [Fact]
    public void FormatConnectionTenantLabel_FallsBackToTenantIdWhenNoNameCached()
    {
        var connection = MakeConnection("1", "alice@contoso.com", tenantId: "guid-123", tenantName: null);
        Assert.Equal("guid-123", W365CliApp.FormatConnectionTenantLabel(connection));
    }

    [Fact]
    public void FormatConnectionTenantLabel_FallsBackToDashWhenNothingKnown()
    {
        var connection = MakeConnection("1", "alice@contoso.com");
        Assert.Equal("-", W365CliApp.FormatConnectionTenantLabel(connection));
    }

    [Fact]
    public void FormatConnectionStatusMarkup_ActiveIsGreenCachedIsGrey()
    {
        Assert.Equal("[green]Active[/]", W365CliApp.FormatConnectionStatusMarkup(MakeConnection("1", "a@b.com", isActive: true)));
        Assert.Equal("[grey]Cached[/]", W365CliApp.FormatConnectionStatusMarkup(MakeConnection("1", "a@b.com", isActive: false)));
    }
}
