using Xunit;

namespace W365Cli.Tests;

public class ProvisioningPolicyTests
{
    private static ProvisioningPolicySummary MakePolicy(
        string name,
        string? provisioningType = "dedicated",
        string? description = null,
        string? imageDisplayName = null)
    {
        return new ProvisioningPolicySummary
        {
            Id = Guid.NewGuid().ToString(),
            DisplayName = name,
            ProvisioningType = provisioningType,
            Description = description,
            ImageDisplayName = imageDisplayName
        };
    }

    [Theory]
    [InlineData("shared", true)]
    [InlineData("sharedByUser", true)]
    [InlineData("sharedByEntraGroup", true)]
    [InlineData("dedicated", false)]
    [InlineData(null, false)]
    public void IsSharedProvisioningPolicy_MatchesKnownSharedTypes(string? provisioningType, bool expected)
    {
        var policy = MakePolicy("Policy", provisioningType);
        Assert.Equal(expected, W365CliApp.IsSharedProvisioningPolicy(policy));
    }

    [Fact]
    public void IsSharedByEntraGroupPolicy_OnlyTrueForThatExactType()
    {
        Assert.True(W365CliApp.IsSharedByEntraGroupPolicy(MakePolicy("A", "sharedByEntraGroup")));
        Assert.False(W365CliApp.IsSharedByEntraGroupPolicy(MakePolicy("A", "sharedByUser")));
        Assert.False(W365CliApp.IsSharedByEntraGroupPolicy(MakePolicy("A", "dedicated")));
    }

    [Fact]
    public void IsSharedByUserPolicy_OnlyTrueForThatExactType()
    {
        Assert.True(W365CliApp.IsSharedByUserPolicy(MakePolicy("A", "sharedByUser")));
        Assert.False(W365CliApp.IsSharedByUserPolicy(MakePolicy("A", "sharedByEntraGroup")));
    }

    [Fact]
    public void FilterProvisioningPolicies_EmptyFilter_ReturnsAllUnchanged()
    {
        var policies = new[] { MakePolicy("Alpha"), MakePolicy("Beta") };
        var result = W365CliApp.FilterProvisioningPolicies(policies, "");
        Assert.Same(policies, result);
    }

    [Fact]
    public void FilterProvisioningPolicies_MatchesDescription()
    {
        var policies = new[]
        {
            MakePolicy("Alpha", description: "Finance team pool"),
            MakePolicy("Beta", description: "Engineering team pool")
        };
        var result = W365CliApp.FilterProvisioningPolicies(policies, "finance");
        Assert.Single(result);
        Assert.Equal("Alpha", result[0].DisplayName);
    }

    [Fact]
    public void FilterProvisioningPolicies_MatchesImageDisplayName()
    {
        var policies = new[]
        {
            MakePolicy("Alpha", imageDisplayName: "Windows 11 Enterprise"),
            MakePolicy("Beta", imageDisplayName: "Windows 10 Enterprise")
        };
        var result = W365CliApp.FilterProvisioningPolicies(policies, "windows 11");
        Assert.Single(result);
        Assert.Equal("Alpha", result[0].DisplayName);
    }

    [Fact]
    public void SortProvisioningPolicies_DefaultSortsByDisplayName()
    {
        var policies = new[] { MakePolicy("Zeta"), MakePolicy("Alpha"), MakePolicy("Mu") };
        var result = W365CliApp.SortProvisioningPolicies(policies, default);
        Assert.Equal(["Alpha", "Mu", "Zeta"], result.Select(p => p.DisplayName));
    }

    [Fact]
    public void SortProvisioningPolicies_ByType_GroupsAndOrdersByType()
    {
        var policies = new[]
        {
            MakePolicy("B-Dedicated", "dedicated"),
            MakePolicy("A-Shared", "shared")
        };
        var result = W365CliApp.SortProvisioningPolicies(policies, W365CliApp.ProvisioningPolicySortMode.Type);
        Assert.Equal(["B-Dedicated", "A-Shared"], result.Select(p => p.DisplayName));
    }

    [Fact]
    public void NextProvisioningPolicySortMode_CyclesThroughAllModes()
    {
        var name = W365CliApp.ProvisioningPolicySortMode.Name;
        var type = W365CliApp.NextProvisioningPolicySortMode(name);
        var image = W365CliApp.NextProvisioningPolicySortMode(type);
        var join = W365CliApp.NextProvisioningPolicySortMode(image);
        var backToName = W365CliApp.NextProvisioningPolicySortMode(join);

        Assert.Equal(W365CliApp.ProvisioningPolicySortMode.Type, type);
        Assert.Equal(W365CliApp.ProvisioningPolicySortMode.Image, image);
        Assert.Equal(W365CliApp.ProvisioningPolicySortMode.Join, join);
        Assert.Equal(W365CliApp.ProvisioningPolicySortMode.Name, backToName);
    }
}
