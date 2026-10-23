using Xunit;

namespace W365Cli.Tests;

/// <summary>
/// Regression coverage for a real bug: Export Markdown Snapshot, Last Sign-In Report, and Client
/// Connection Type Report were each added to ShowReportsAsync's flat PromptChoice list (Reports
/// menu) but never added to GetMainMenuChoices()'s "Reports" tree node -- the list that actually
/// drives the sidebar/dashboard navigation tree and the Ctrl+K command palette. The two lists were
/// hand-maintained in parallel with no shared source, so a report could be fully wired up and
/// reachable via the plain Reports menu while silently invisible everywhere else. Both now build
/// from the single ReportMenuItems array; this test asserts they can never drift apart again.
/// </summary>
public class ReportMenuConsistencyTests
{
    [Fact]
    public void ReportsNode_InMainMenuTree_ContainsEveryReportMenuItem()
    {
        var app = new W365CliApp();
        var reportsNode = Assert.Single(app.GetMainMenuChoices(), choice => choice.Key == "Reports");
        Assert.NotNull(reportsNode.Children);

        var treeTitles = reportsNode.Children!.Select(child => child.Title).ToHashSet(StringComparer.Ordinal);
        var sourceTitles = W365CliApp.ReportMenuItems.Select(item => item.Title).ToHashSet(StringComparer.Ordinal);

        Assert.Equal(sourceTitles, treeTitles);
    }

    [Fact]
    public void ReportMenuItems_HasNoDuplicateTitles()
    {
        var titles = W365CliApp.ReportMenuItems.Select(item => item.Title).ToArray();
        Assert.Equal(titles.Distinct(StringComparer.Ordinal).Count(), titles.Length);
    }
}
