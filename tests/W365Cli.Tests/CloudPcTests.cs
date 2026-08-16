using Xunit;

namespace W365Cli.Tests;

public class CloudPcTests
{
    private static CloudPcSummary MakePc(
        string name,
        string? status = "provisioned",
        string? provisioningType = "dedicated",
        string? upn = null,
        string? servicePlan = null,
        string? realTimeSignInStatus = null,
        CloudPcSharedDeviceDetail? sharedDeviceDetail = null,
        CloudPcConnectivityResult? connectivityResult = null)
    {
        return new CloudPcSummary
        {
            Id = Guid.NewGuid().ToString(),
            DisplayName = name,
            Status = status,
            ProvisioningType = provisioningType,
            UserPrincipalName = upn,
            ServicePlanName = servicePlan,
            RealTimeSignInStatus = realTimeSignInStatus,
            SharedDeviceDetail = sharedDeviceDetail,
            ConnectivityResult = connectivityResult
        };
    }

    [Fact]
    public void FilterCloudPcs_EmptyFilter_ReturnsAllUnchanged()
    {
        var pcs = new[] { MakePc("Alpha"), MakePc("Beta") };
        var result = W365CliApp.FilterCloudPcs(pcs, "");
        Assert.Same(pcs, result);
    }

    [Theory]
    [InlineData("alpha")]
    [InlineData("ALPHA")]
    public void FilterCloudPcs_MatchesNameCaseInsensitive(string filter)
    {
        var pcs = new[] { MakePc("Alpha"), MakePc("Beta") };
        var result = W365CliApp.FilterCloudPcs(pcs, filter);
        Assert.Single(result);
        Assert.Equal("Alpha", result[0].Name);
    }

    [Fact]
    public void FilterCloudPcs_MatchesUserPrincipalName()
    {
        var pcs = new[]
        {
            MakePc("Alpha", upn: "jane@contoso.com"),
            MakePc("Beta", upn: "john@contoso.com")
        };
        var result = W365CliApp.FilterCloudPcs(pcs, "jane");
        Assert.Single(result);
        Assert.Equal("Alpha", result[0].Name);
    }

    [Fact]
    public void FilterCloudPcs_MatchesServicePlanName()
    {
        var pcs = new[]
        {
            MakePc("Alpha", servicePlan: "Windows 365 Enterprise 4 vCPU"),
            MakePc("Beta", servicePlan: "Windows 365 Frontline")
        };
        var result = W365CliApp.FilterCloudPcs(pcs, "frontline");
        Assert.Single(result);
        Assert.Equal("Beta", result[0].Name);
    }

    [Fact]
    public void SortCloudPcs_DefaultSortsByNameOrdinalIgnoreCase()
    {
        var pcs = new[] { MakePc("Zeta"), MakePc("alpha"), MakePc("Mu") };
        var result = W365CliApp.SortCloudPcs(pcs, default);
        Assert.Equal(["alpha", "Mu", "Zeta"], result.Select(pc => pc.Name));
    }

    [Fact]
    public void SortCloudPcs_ByUser_SortsByEffectiveUpnThenName()
    {
        var pcs = new[]
        {
            MakePc("B-PC", upn: "zed@contoso.com"),
            MakePc("A-PC", upn: "amy@contoso.com")
        };
        var result = W365CliApp.SortCloudPcs(pcs, W365CliApp.CloudPcSortMode.User);
        Assert.Equal(["A-PC", "B-PC"], result.Select(pc => pc.Name));
    }

    [Fact]
    public void NextCloudPcSortMode_CyclesThroughAllModesAndWrapsAround()
    {
        var name = W365CliApp.CloudPcSortMode.Name;
        var status = W365CliApp.NextCloudPcSortMode(name);
        var user = W365CliApp.NextCloudPcSortMode(status);
        var servicePlan = W365CliApp.NextCloudPcSortMode(user);
        var backToName = W365CliApp.NextCloudPcSortMode(servicePlan);

        Assert.Equal(W365CliApp.CloudPcSortMode.Status, status);
        Assert.Equal(W365CliApp.CloudPcSortMode.User, user);
        Assert.Equal(W365CliApp.CloudPcSortMode.ServicePlan, servicePlan);
        Assert.Equal(W365CliApp.CloudPcSortMode.Name, backToName);
    }

    [Fact]
    public void FormatCloudPcSortMode_ReturnsHumanReadableLabels()
    {
        Assert.Equal("name", W365CliApp.FormatCloudPcSortMode(W365CliApp.CloudPcSortMode.Name));
        Assert.Equal("status", W365CliApp.FormatCloudPcSortMode(W365CliApp.CloudPcSortMode.Status));
        Assert.Equal("user", W365CliApp.FormatCloudPcSortMode(W365CliApp.CloudPcSortMode.User));
        Assert.Equal("service plan", W365CliApp.FormatCloudPcSortMode(W365CliApp.CloudPcSortMode.ServicePlan));
    }

    // --- In-use status: this logic was heavily debugged this session (real-time sign-in status
    // vs. the unreliable bulk connectivityResult, Enterprise vs. Flex session-start handling,
    // Unavailable-vs-Available for provisioned vs. notProvisioned) -- high regression risk area.

    [Fact]
    public void GetNormalizedInUseStatus_NotSignedIn_ReturnsAvailable()
    {
        var pc = MakePc("Alpha", realTimeSignInStatus: "NotSignedIn");
        Assert.Equal("available", W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void GetNormalizedInUseStatus_SignedIn_ReturnsInUse()
    {
        var pc = MakePc("Alpha", realTimeSignInStatus: "SignedIn");
        Assert.Equal("inUse", W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void GetNormalizedInUseStatus_UnavailableAndNotProvisioned_ReturnsUnavailable()
    {
        var pc = MakePc("Alpha", status: "notProvisioned", realTimeSignInStatus: "Unavailable");
        Assert.Equal("unavailable", W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void GetNormalizedInUseStatus_UnavailableButProvisioned_ReturnsAvailable()
    {
        var pc = MakePc("Alpha", status: "provisioned", realTimeSignInStatus: "Unavailable");
        Assert.Equal("available", W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void GetNormalizedInUseStatus_FallsBackToConnectivityResult_WhenNoRealTimeStatus()
    {
        var pc = MakePc("Alpha", connectivityResult: new CloudPcConnectivityResult { Status = "available" });
        Assert.Equal("available", W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void GetNormalizedInUseStatus_NoDataAtAll_ReturnsNull()
    {
        var pc = MakePc("Alpha");
        Assert.Null(W365CliApp.GetNormalizedInUseStatus(pc));
    }

    [Fact]
    public void FormatInUsePlain_InUseWithSessionStart_ShowsSinceTime()
    {
        var start = new DateTimeOffset(2024, 6, 1, 13, 0, 0, TimeSpan.Zero);
        var pc = MakePc("Alpha", realTimeSignInStatus: "SignedIn", sharedDeviceDetail: new CloudPcSharedDeviceDetail
        {
            SessionStartDateTime = start
        });

        var result = W365CliApp.FormatInUsePlain(pc);
        Assert.StartsWith("In use since ", result, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatInUsePlain_InUseWithoutSessionStart_ShowsGenericInUse()
    {
        // Enterprise (dedicated) Cloud PCs have no SharedDeviceDetail-based session start; the
        // known-inaccurate RealTimeLastActiveTime heartbeat must never be used as a substitute.
        var pc = MakePc("Alpha", realTimeSignInStatus: "SignedIn");
        Assert.Equal("In use", W365CliApp.FormatInUsePlain(pc));
    }

    [Fact]
    public void FormatInUsePlain_NoStatus_ReturnsDash()
    {
        var pc = MakePc("Alpha");
        Assert.Equal("-", W365CliApp.FormatInUsePlain(pc));
    }

    [Fact]
    public void FormatInUsePlain_Available_ReturnsAvailable()
    {
        var pc = MakePc("Alpha", realTimeSignInStatus: "NotSignedIn");
        Assert.Equal("Available", W365CliApp.FormatInUsePlain(pc));
    }

    [Fact]
    public void FormatInUsePlain_Unavailable_ReturnsUnavailable()
    {
        var pc = MakePc("Alpha", status: "notProvisioned", realTimeSignInStatus: "Unavailable");
        Assert.Equal("Unavailable", W365CliApp.FormatInUsePlain(pc));
    }

    [Fact]
    public void FormatInUseMarkup_InUse_UsesYellow()
    {
        var pc = MakePc("Alpha", realTimeSignInStatus: "SignedIn");
        Assert.Contains("[yellow]", W365CliApp.FormatInUseMarkup(pc), StringComparison.Ordinal);
    }

    [Fact]
    public void FormatInUseMarkup_Available_UsesGreen()
    {
        var pc = MakePc("Alpha", realTimeSignInStatus: "NotSignedIn");
        Assert.Contains("[green]", W365CliApp.FormatInUseMarkup(pc), StringComparison.Ordinal);
    }

    [Fact]
    public void FormatInUseMarkup_Unavailable_UsesGrey()
    {
        var pc = MakePc("Alpha", status: "notProvisioned", realTimeSignInStatus: "Unavailable");
        Assert.Contains("[grey]", W365CliApp.FormatInUseMarkup(pc), StringComparison.Ordinal);
    }
}
