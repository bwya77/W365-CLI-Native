using System.Text.Json;
using Xunit;

namespace W365Cli.Tests;

public class ReportParsingTests
{
    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement;

    [Fact]
    public void ParseReportRows_ConnectionHistory_ProducesOneRowPerValueRow()
    {
        var report = Parse(ReportJsonFixtures.ConnectionHistoryReport);
        var rows = W365GraphClient.ParseReportRows(report, "UPN", "ClientOS", "TransportType");

        Assert.Equal(2, rows.Count);
        Assert.Equal("user1@contoso.com", rows[0].Fields["UPN"]);
        Assert.Equal("2024-06-01T13:00:00Z", rows[0].Fields["SessionBeginTime"]);
        Assert.Equal("Windows 11", rows[0].Fields["ClientOS"]);
    }

    [Fact]
    public void ParseReportRows_SummaryUsesRequestedFieldsInOrder()
    {
        var report = Parse(ReportJsonFixtures.ConnectionHistoryReport);
        var rows = W365GraphClient.ParseReportRows(report, "UPN", "ClientOS");

        Assert.Equal("user1@contoso.com | Windows 11", rows[0].Summary);
    }

    [Fact]
    public void ParseReportRows_EmptyReport_ReturnsNoRows()
    {
        var report = Parse(ReportJsonFixtures.EmptyReport);
        var rows = W365GraphClient.ParseReportRows(report, "UPN");
        Assert.Empty(rows);
    }

    [Fact]
    public void ParseReportRows_MalformedReport_ReturnsNoRowsInsteadOfThrowing()
    {
        var report = Parse(ReportJsonFixtures.MalformedReport);
        var rows = W365GraphClient.ParseReportRows(report, "UPN");
        Assert.Empty(rows);
    }

    [Fact]
    public void ParseReportRowsAdaptive_DropsGuidLikeAndNoiseColumnsFromSummary()
    {
        // Real regression scenario: frontlineLicenseHourlyUsageReport's useful columns
        // (LicenseCount, ClaimedLicenseCount, SkuLicenseCount) aren't part of any fixed summary
        // field list, so ParseReportRowsAdaptive must surface them generically while dropping the
        // synthetic UniqueId key and IngestedTimestamp/GUID noise.
        var report = Parse(ReportJsonFixtures.FrontlineLicenseHourlyUsageReport);
        var rows = W365GraphClient.ParseReportRowsAdaptive(report);

        Assert.Equal(2, rows.Count);
        var firstSummary = rows[0].Summary;
        Assert.Contains("LicenseCount: 10", firstSummary, StringComparison.Ordinal);
        Assert.Contains("ClaimedLicenseCount: 7", firstSummary, StringComparison.Ordinal);
        Assert.DoesNotContain("UniqueId", firstSummary, StringComparison.Ordinal);
        Assert.DoesNotContain("33333333-3333-3333-3333-333333333333", firstSummary, StringComparison.Ordinal);
    }

    [Fact]
    public void ParseReportRowsAdaptive_TitleFallsBackToDisplayName()
    {
        var report = Parse(ReportJsonFixtures.FrontlineLicenseHourlyUsageReport);
        var rows = W365GraphClient.ParseReportRowsAdaptive(report);
        Assert.Equal("Frontline Flex Pool", rows[0].Title);
    }

    [Fact]
    public void GetReportColumnName_UsesColumnPropertyFromSchemaObject()
    {
        var schemaItem = Parse("""{ "Column": "UPN" }""");
        Assert.Equal("UPN", W365GraphClient.GetReportColumnName(schemaItem));
    }

    [Fact]
    public void GetReportColumnName_HandlesPlainStringSchemaEntries()
    {
        var schemaItem = Parse("\"UPN\"");
        Assert.Equal("UPN", W365GraphClient.GetReportColumnName(schemaItem));
    }
}
