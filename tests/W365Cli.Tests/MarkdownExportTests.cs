using System.Text;
using Xunit;

namespace W365Cli.Tests;

public class MarkdownExportTests
{
    [Theory]
    [InlineData("plain", "plain")]
    [InlineData(null, "-")]
    [InlineData("", "-")]
    [InlineData("a|b", "a\\|b")]
    [InlineData("line1\r\nline2", "line1 line2")]
    public void MarkdownEscape_HandlesNullPipesAndNewlines(string? input, string expected)
    {
        Assert.Equal(expected, W365CliApp.MarkdownEscape(input));
    }

    [Fact]
    public void AppendMarkdownTable_WritesHeaderSeparatorAndRows()
    {
        var sb = new StringBuilder();
        W365CliApp.AppendMarkdownTable(
            sb,
            ["Name", "Status"],
            [["Alpha", "provisioned"], ["Beta", "failed"]]);

        var output = sb.ToString();
        Assert.Contains("| Name | Status |", output, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", output, StringComparison.Ordinal);
        Assert.Contains("| Alpha | provisioned |", output, StringComparison.Ordinal);
        Assert.Contains("| Beta | failed |", output, StringComparison.Ordinal);
    }

    [Fact]
    public void AppendCloudPcSection_EmptyList_ShowsPlaceholderMessage()
    {
        var sb = new StringBuilder();
        W365CliApp.AppendCloudPcSection(sb, []);
        Assert.Contains("_No Cloud PCs were returned._", sb.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void AppendCloudPcSection_GroupsByStatusAndType()
    {
        var cloudPcs = new[]
        {
            new CloudPcSummary { Id = "1", DisplayName = "Alpha", Status = "provisioned", ProvisioningType = "dedicated" },
            new CloudPcSummary { Id = "2", DisplayName = "Beta", Status = "failed", ProvisioningType = "shared" }
        };

        var sb = new StringBuilder();
        W365CliApp.AppendCloudPcSection(sb, cloudPcs);
        var output = sb.ToString();

        Assert.Contains("Total: **2**", output, StringComparison.Ordinal);
        Assert.Contains("failed: 1", output, StringComparison.Ordinal);
        Assert.Contains("provisioned: 1", output, StringComparison.Ordinal);
        Assert.Contains("dedicated: 1", output, StringComparison.Ordinal);
        Assert.Contains("shared: 1", output, StringComparison.Ordinal);
        Assert.Contains("| Alpha |", output, StringComparison.Ordinal);
        Assert.Contains("| Beta |", output, StringComparison.Ordinal);
    }

    [Fact]
    public void AppendProvisioningPolicySection_EmptyList_ShowsPlaceholderMessage()
    {
        var sb = new StringBuilder();
        W365CliApp.AppendProvisioningPolicySection(sb, []);
        Assert.Contains("_No provisioning policies were returned._", sb.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void AppendLicensingSection_WithError_ShowsErrorBlockQuote()
    {
        var sb = new StringBuilder();
        W365CliApp.AppendLicensingSection(sb, [], "Access denied reading subscribedSkus");
        var output = sb.ToString();
        Assert.Contains("Licensing data could not be loaded", output, StringComparison.Ordinal);
        Assert.Contains("> Access denied reading subscribedSkus", output, StringComparison.Ordinal);
    }

    [Fact]
    public void AppendLicensingSection_NoError_NoItems_ShowsPlaceholder()
    {
        var sb = new StringBuilder();
        W365CliApp.AppendLicensingSection(sb, [], null);
        Assert.Contains("_No Windows 365 license SKUs were detected._", sb.ToString(), StringComparison.Ordinal);
    }
}
