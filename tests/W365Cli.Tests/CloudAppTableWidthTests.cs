using Xunit;

namespace W365Cli.Tests;

public class CloudAppTableWidthTests
{
    // Same class of regression this table shares with Cloud PCs: at a terminal width right around
    // where "show dates" turns on (Console.WindowWidth >= 150) combined with the side-by-side
    // "Selected Cloud App" panel, the previous flat "Math.Max(48, ...)" floor could push the total
    // requested table width past the real terminal width, and Spectre would silently shrink
    // whichever column it picked below what NoWrap could truncate cleanly -- wrapping cell text
    // character-by-character instead of cropping it with "...".

    private static int ComputeTotalTableWidth(
        (int Status, int Name, int Publisher, int Published, int Added) widths,
        bool showPublisher,
        bool showDates,
        bool sideBySide)
    {
        var columnCount = 3 + (showPublisher ? 1 : 0) + (showDates ? 2 : 0);
        var overhead = (3 * columnCount) + 1;
        var sidePanelCost = sideBySide ? 40 : 0;
        return overhead + sidePanelCost + 1 /* selector */ + widths.Status + widths.Name + widths.Publisher + widths.Published + widths.Added;
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void GetMinimumSideBySideCloudAppTableWidth_MatchesTheTrueFloorOfAllColumnMinimums(bool showPublisher, bool showDates)
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudAppTableWidth(showPublisher, showDates);
        var widths = W365CliApp.GetCloudAppWidths(showPublisher, showDates, sideBySide: true, minimum);
        var total = ComputeTotalTableWidth(widths, showPublisher, showDates, sideBySide: true);

        Assert.True(total <= minimum, $"Computed table width {total} exceeds the claimed minimum {minimum}.");
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void GetCloudAppWidths_AtOrAboveTheSideBySideMinimum_EveryColumnStaysAtOrAboveItsMinimumAndTableFits(bool showPublisher, bool showDates)
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudAppTableWidth(showPublisher, showDates);

        foreach (var windowWidth in new[] { minimum, minimum + 5, minimum + 20, minimum + 100 })
        {
            var widths = W365CliApp.GetCloudAppWidths(showPublisher, showDates, sideBySide: true, windowWidth);

            Assert.True(widths.Status >= 10, $"[{windowWidth}] Status width {widths.Status} fell below its minimum.");
            Assert.True(widths.Name >= 16, $"[{windowWidth}] Name width {widths.Name} fell below its minimum.");
            if (showPublisher)
            {
                Assert.True(widths.Publisher >= 14, $"[{windowWidth}] Publisher width {widths.Publisher} fell below its minimum.");
            }

            if (showDates)
            {
                Assert.True(widths.Published >= 12, $"[{windowWidth}] Published width {widths.Published} fell below its minimum.");
                Assert.True(widths.Added >= 12, $"[{windowWidth}] Added width {widths.Added} fell below its minimum.");
            }

            var total = ComputeTotalTableWidth(widths, showPublisher, showDates, sideBySide: true);
            Assert.True(total <= windowWidth, $"Computed table width {total} exceeds window width {windowWidth}.");
        }
    }

    [Fact]
    public void GetCloudAppWidths_RealWorld150ColumnRegression_DoesNotOverflow()
    {
        // The specific real-world width (Console.WindowWidth == 150) where showDates flips on
        // while the side-by-side panel is already showing (>=125) -- this combination is exactly
        // what triggered the reported bug before the fix.
        var widths = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: true, sideBySide: true, windowWidth: 150);
        var total = ComputeTotalTableWidth(widths, showPublisher: true, showDates: true, sideBySide: true);

        Assert.True(total <= 150, $"Computed table width {total} exceeds window width 150.");
    }

    [Fact]
    public void GetCloudAppWidths_DatesAndStatusShrinkBeforeNameWhenSpaceIsTight()
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudAppTableWidth(showPublisher: true, showDates: true);
        var narrow = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: true, sideBySide: true, minimum);
        var wide = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: true, sideBySide: true, windowWidth: 250);

        Assert.True(narrow.Published <= wide.Published);
        Assert.True(narrow.Added <= wide.Added);
        Assert.True(narrow.Name >= 16);
    }

    [Fact]
    public void GetCloudAppWidths_PlentyOfRoom_CapsNameAndPublisherInsteadOfGrowingUnbounded()
    {
        var widths = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: false, sideBySide: true, windowWidth: 500);

        Assert.Equal(60, widths.Name);
        Assert.Equal(40, widths.Publisher);
    }

    [Fact]
    public void GetCloudAppWidths_NotSideBySide_HasMoreRoomThanSideBySideAtSameWindowWidth()
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudAppTableWidth(showPublisher: true, showDates: true);
        var sideBySide = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: true, sideBySide: true, minimum);
        var standalone = W365CliApp.GetCloudAppWidths(showPublisher: true, showDates: true, sideBySide: false, minimum);

        Assert.True(standalone.Name >= sideBySide.Name);
    }

    // Regression coverage for a real reported visual issue: Publisher is "-" (unset) for most
    // built-in Microsoft/Windows apps, but the column always rendered at its full allocated width
    // (up to 40 chars) regardless -- a big, mostly-empty column that made the table look lopsided
    // even though nothing was being truncated. GetContentAwareColumnWidth lets the caller shrink a
    // column down to what the actually-visible values need instead of always using the full budget.

    [Fact]
    public void GetContentAwareColumnWidth_AllBlankValues_ShrinksToMinimum()
    {
        var width = W365CliApp.GetContentAwareColumnWidth(new string?[] { null, "", "   " }, minWidth: 14, maxWidth: 40);
        Assert.Equal(14, width);
    }

    [Fact]
    public void GetContentAwareColumnWidth_UsesLongestActualValue()
    {
        var width = W365CliApp.GetContentAwareColumnWidth(new string?[] { "Microsoft", null, "Contoso Ltd" }, minWidth: 14, maxWidth: 40);
        Assert.Equal(14, width); // "Contoso Ltd" (11 chars) and "Microsoft" (9 chars) both fall below the minimum
    }

    [Fact]
    public void GetContentAwareColumnWidth_LongestValueBetweenMinAndMax_UsesExactLength()
    {
        var width = W365CliApp.GetContentAwareColumnWidth(new string?[] { "A Reasonably Long Publisher Name" }, minWidth: 14, maxWidth: 40);
        Assert.Equal(32, width);
    }

    [Fact]
    public void GetContentAwareColumnWidth_LongestValueExceedsMax_ClampsToMax()
    {
        var width = W365CliApp.GetContentAwareColumnWidth(new string?[] { new string('X', 100) }, minWidth: 14, maxWidth: 40);
        Assert.Equal(40, width);
    }

    [Fact]
    public void GetContentAwareColumnWidth_EmptyCollection_ReturnsMinimum()
    {
        var width = W365CliApp.GetContentAwareColumnWidth(Array.Empty<string?>(), minWidth: 14, maxWidth: 40);
        Assert.Equal(14, width);
    }
}
