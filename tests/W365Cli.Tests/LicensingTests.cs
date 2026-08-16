using Xunit;

namespace W365Cli.Tests;

public class LicensingTests
{
    [Theory]
    [InlineData("2 vCPU / 8 GB RAM / 128 GB storage", "2/8/128")]
    [InlineData("4 vCPU, 16GB RAM, 256GB Storage", "4/16/256")]
    [InlineData("CPC_E_4C_16GB_128GB", "4/16/128")]
    [InlineData("no numbers here", null)]
    public void GetPlanKey_ExtractsCpuRamStorageFromDisplayOrSkuText(string input, string? expected)
    {
        Assert.Equal(expected, W365CliApp.GetPlanKey(input));
    }

    [Fact]
    public void FormatPlanKey_FormatsThreePartKeyAsHumanReadable()
    {
        Assert.Equal("4vCPU/16GB/128GB", W365CliApp.FormatPlanKey("4/16/128"));
    }

    [Fact]
    public void FormatPlanKey_ReturnsInputUnchangedWhenNotThreeParts()
    {
        Assert.Equal("unknown", W365CliApp.FormatPlanKey("unknown"));
    }

    [Theory]
    [InlineData("Windows 365 Reserve", true)]
    [InlineData("CPC_R_something", true)]
    [InlineData("WINDOWS_365_R_2C_8GB", true)]
    [InlineData("Windows 365 Enterprise", false)]
    public void IsReserveText_DetectsReserveMarkers(string value, bool expected)
    {
        Assert.Equal(expected, W365CliApp.IsReserveText(value));
    }

    [Fact]
    public void GetWindows365LicenseInfo_NonWindows365Sku_ReturnsNull()
    {
        var sku = new SubscribedSku { SkuPartNumber = "OFFICE_365_E3" };
        Assert.Null(W365CliApp.GetWindows365LicenseInfo(sku));
    }

    [Fact]
    public void GetWindows365LicenseInfo_AddOnSku_ReturnsNull()
    {
        var sku = new SubscribedSku { SkuPartNumber = "CPC_ADD_ON" };
        Assert.Null(W365CliApp.GetWindows365LicenseInfo(sku));
    }

    [Fact]
    public void GetWindows365LicenseInfo_FlexSku_ClassifiedAsFlexFamily()
    {
        var sku = new SubscribedSku { SkuPartNumber = "CPC_F_4C_16GB_128GB" };
        var info = W365CliApp.GetWindows365LicenseInfo(sku);
        Assert.NotNull(info);
        Assert.Equal("Flex", info!.Family);
    }

    [Fact]
    public void GetWindows365LicenseInfo_EnterpriseSku_ClassifiedAsEnterpriseFamily()
    {
        var sku = new SubscribedSku { SkuPartNumber = "CPC_E_4C_16GB_128GB" };
        var info = W365CliApp.GetWindows365LicenseInfo(sku);
        Assert.NotNull(info);
        Assert.Equal("Enterprise", info!.Family);
    }

    [Fact]
    public void GetWindows365LicenseInfo_ReserveSku_ClassifiedAsReserveFamily()
    {
        var sku = new SubscribedSku { SkuPartNumber = "CPC_R_4C_16GB_128GB" };
        var info = W365CliApp.GetWindows365LicenseInfo(sku);
        Assert.NotNull(info);
        Assert.Equal("Reserve", info!.Family);
    }

    [Fact]
    public void BuildLicenseOverview_EnterpriseSku_CountsPurchasedAssignedAndMatchingCloudPcs()
    {
        var skus = new[]
        {
            new SubscribedSku
            {
                SkuPartNumber = "CPC_E_4C_16GB_128GB",
                PrepaidUnits = new SubscribedSkuPrepaidUnits { Enabled = 10 },
                ConsumedUnits = 3,
                ServicePlans = []
            }
        };
        var cloudPcs = new[]
        {
            new CloudPcSummary
            {
                Id = "1",
                DisplayName = "CPC-1",
                ServicePlanName = "Windows 365 Enterprise 4 vCPU/16 GB RAM/128 GB Storage",
                ProvisioningType = "dedicated"
            },
            new CloudPcSummary
            {
                Id = "2",
                DisplayName = "CPC-2",
                ServicePlanName = "Windows 365 Frontline 2 vCPU/8 GB RAM/64 GB Storage",
                ProvisioningType = "shared"
            }
        };

        var overview = W365CliApp.BuildLicenseOverview(skus, cloudPcs, []);

        Assert.Single(overview);
        var item = overview[0];
        Assert.Equal(10, item.Purchased);
        Assert.Equal(3, item.Assigned);
        Assert.Equal(1, item.CloudPcCount);
        Assert.Contains(item.CloudPcs, pc => pc.Id == "1");
        Assert.DoesNotContain(item.CloudPcs, pc => pc.Id == "2");
    }

    [Fact]
    public void BuildLicenseOverview_FlexSku_SplitsDedicatedAndSharedUnitUsage()
    {
        var skus = new[]
        {
            new SubscribedSku
            {
                SkuPartNumber = "CPC_F_2C_8GB_64GB",
                PrepaidUnits = new SubscribedSkuPrepaidUnits { Enabled = 5 },
                ConsumedUnits = 0,
                ServicePlans = []
            }
        };
        // 3 dedicated Flex Cloud PCs use ceil(3/3)=1 license unit; 2 shared pool Cloud PCs use 2
        // license units (1 each) -- exercising the Flex license-math engine's core rule that one
        // license unit covers either 1 shared Cloud PC or up to 3 dedicated ones.
        var cloudPcs = new List<CloudPcSummary>();
        for (var i = 0; i < 3; i++)
        {
            cloudPcs.Add(new CloudPcSummary
            {
                Id = $"dedicated-{i}",
                DisplayName = $"CPC-D-{i}",
                ServicePlanName = "Windows 365 Frontline 2 vCPU/8 GB RAM/64 GB Storage",
                ProvisioningType = "dedicated",
                ProvisioningPolicyName = "Flex Dedicated Policy"
            });
        }
        for (var i = 0; i < 2; i++)
        {
            cloudPcs.Add(new CloudPcSummary
            {
                Id = $"shared-{i}",
                DisplayName = $"CPC-S-{i}",
                ServicePlanName = "Windows 365 Frontline 2 vCPU/8 GB RAM/64 GB Storage",
                ProvisioningType = "shared",
                ProvisioningPolicyName = "Flex Shared Policy"
            });
        }

        var overview = W365CliApp.BuildLicenseOverview(skus, cloudPcs, []);

        Assert.Single(overview);
        var item = overview[0];
        Assert.Equal(5, item.Purchased);
        Assert.Equal(5, item.CloudPcCount);
        Assert.Equal(3, item.DedicatedCloudPcCount);
        Assert.Equal(2, item.SharedCloudPcCount);
        Assert.Equal(1, item.DedicatedUnitsUsed);
        Assert.Equal(2, item.SharedUnitsUsed);
        Assert.Equal(3, item.LicenseUnitsUsed);
        Assert.Equal(2, item.LicenseUnitsLeft);
    }

    [Fact]
    public void BuildLicenseOverview_NoWindows365Skus_ReturnsEmpty()
    {
        var skus = new[] { new SubscribedSku { SkuPartNumber = "OFFICE_365_E3" } };
        var overview = W365CliApp.BuildLicenseOverview(skus, [], []);
        Assert.Empty(overview);
    }
}
