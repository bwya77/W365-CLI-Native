using Xunit;

namespace W365Cli.Tests;

public class GraphRowTests
{
    private static GraphTableRow MakeRow(string title, string summary, Dictionary<string, string>? fields = null)
    {
        return new GraphTableRow(title, summary, fields ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void FilterGraphRows_EmptyFilter_ReturnsAllUnchanged()
    {
        var rows = new[] { MakeRow("Alpha", "summary-a"), MakeRow("Beta", "summary-b") };
        var result = W365CliApp.FilterGraphRows(rows, "");
        Assert.Same(rows, result);
    }

    [Fact]
    public void FilterGraphRows_MatchesTitleSummaryOrFieldValue()
    {
        var rows = new[]
        {
            MakeRow("Alpha", "nothing interesting"),
            MakeRow("Beta", "nothing interesting", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["ClientOS"] = "macOS"
            })
        };

        var result = W365CliApp.FilterGraphRows(rows, "macos");
        Assert.Single(result);
        Assert.Equal("Beta", result[0].Title);
    }

    [Fact]
    public void SortGraphRows_TitleAscending_OrdersByTitle()
    {
        var rows = new[] { MakeRow("Zeta", "z"), MakeRow("Alpha", "a") };
        var result = W365CliApp.SortGraphRows(rows, W365CliApp.GraphRowSortMode.TitleAscending);
        Assert.Equal(["Alpha", "Zeta"], result.Select(r => r.Title));
    }

    [Fact]
    public void SortGraphRows_TitleDescending_OrdersReverse()
    {
        var rows = new[] { MakeRow("Alpha", "a"), MakeRow("Zeta", "z") };
        var result = W365CliApp.SortGraphRows(rows, W365CliApp.GraphRowSortMode.TitleDescending);
        Assert.Equal(["Zeta", "Alpha"], result.Select(r => r.Title));
    }

    [Fact]
    public void SortGraphRows_None_PreservesOriginalOrder()
    {
        var rows = new[] { MakeRow("Zeta", "z"), MakeRow("Alpha", "a") };
        var result = W365CliApp.SortGraphRows(rows, W365CliApp.GraphRowSortMode.None);
        Assert.Equal(["Zeta", "Alpha"], result.Select(r => r.Title));
    }

    [Fact]
    public void NextSortMode_CyclesThroughAllFourModesAndWrapsToNone()
    {
        var none = W365CliApp.GraphRowSortMode.None;
        var titleAsc = W365CliApp.NextSortMode(none);
        var titleDesc = W365CliApp.NextSortMode(titleAsc);
        var summaryAsc = W365CliApp.NextSortMode(titleDesc);
        var summaryDesc = W365CliApp.NextSortMode(summaryAsc);
        var backToNone = W365CliApp.NextSortMode(summaryDesc);

        Assert.Equal(W365CliApp.GraphRowSortMode.TitleAscending, titleAsc);
        Assert.Equal(W365CliApp.GraphRowSortMode.TitleDescending, titleDesc);
        Assert.Equal(W365CliApp.GraphRowSortMode.SummaryAscending, summaryAsc);
        Assert.Equal(W365CliApp.GraphRowSortMode.SummaryDescending, summaryDesc);
        Assert.Equal(W365CliApp.GraphRowSortMode.None, backToNone);
    }
}
