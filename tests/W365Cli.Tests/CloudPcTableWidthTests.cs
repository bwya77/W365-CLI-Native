using Xunit;

namespace W365Cli.Tests;

public class CloudPcTableWidthTests
{
    // Regression coverage for a real reported bug: at a narrow-but-still-"wide enough for the side
    // panel" terminal width, Status/Type stayed fixed at their preferred sizes while Name/User were
    // floored at a minimum that still didn't fit -- so the total requested table width exceeded the
    // real terminal width, and Spectre silently shrank whatever column it chose below what NoWrap
    // could truncate cleanly, wrapping cell text character-by-character (e.g. "sharedByUser"
    // rendering as "sh/ar/ed/U.."). These tests assert every column comes back at or above its
    // documented minimum, and that the whole table actually fits inside the given window width, for
    // a range of realistic narrow terminal widths.

    private static int ComputeTotalTableWidth(
        (int Name, int Status, int Type, int User, int ServicePlan, int InUse) widths,
        bool showInUse,
        bool showUser,
        bool showServicePlan,
        bool sideBySide)
    {
        var columnCount = 4 + (showInUse ? 1 : 0) + (showUser ? 1 : 0) + (showServicePlan ? 1 : 0);
        var overhead = (3 * columnCount) + 1;
        var sidePanelCost = sideBySide ? 40 : 0;
        return overhead + sidePanelCost + 1 /* selector */ + widths.Status + widths.Type + widths.InUse + widths.Name + widths.User + widths.ServicePlan;
    }

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, true)]
    public void GetMinimumSideBySideCloudPcTableWidth_MatchesTheTrueFloorOfAllColumnMinimums(bool showInUse, bool showUser, bool showServicePlan)
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudPcTableWidth(showInUse, showUser, showServicePlan);
        var widths = W365CliApp.GetCloudPcWidths(showInUse, showUser, showServicePlan, sideBySide: true, minimum);
        var total = ComputeTotalTableWidth(widths, showInUse, showUser, showServicePlan, sideBySide: true);

        Assert.Equal(minimum, total);
    }

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, true)]
    public void GetCloudPcWidths_AtOrAboveTheSideBySideMinimum_EveryColumnStaysAtOrAboveItsMinimumAndTableFits(bool showInUse, bool showUser, bool showServicePlan)
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudPcTableWidth(showInUse, showUser, showServicePlan);

        foreach (var windowWidth in new[] { minimum, minimum + 5, minimum + 20, minimum + 100 })
        {
            var widths = W365CliApp.GetCloudPcWidths(showInUse, showUser, showServicePlan, sideBySide: true, windowWidth);

            Assert.True(widths.Status >= 12, $"[{windowWidth}] Status width {widths.Status} fell below its minimum.");
            Assert.True(widths.Type >= 9, $"[{windowWidth}] Type width {widths.Type} fell below its minimum.");
            Assert.True(widths.Name >= 16, $"[{windowWidth}] Name width {widths.Name} fell below its minimum.");
            if (showInUse)
            {
                Assert.True(widths.InUse >= 14, $"[{windowWidth}] In use width {widths.InUse} fell below its minimum.");
            }

            if (showUser)
            {
                Assert.True(widths.User >= 16, $"[{windowWidth}] User width {widths.User} fell below its minimum.");
            }

            var total = ComputeTotalTableWidth(widths, showInUse, showUser, showServicePlan, sideBySide: true);
            Assert.True(total <= windowWidth, $"Computed table width {total} exceeds window width {windowWidth}.");
        }
    }

    [Fact]
    public void GetCloudPcWidths_TypeShrinksBeforeNameWhenSpaceIsTight()
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudPcTableWidth(showInUse: true, showUser: true, showServicePlan: false);
        var narrow = W365CliApp.GetCloudPcWidths(showInUse: true, showUser: true, showServicePlan: false, sideBySide: true, minimum);
        var wide = W365CliApp.GetCloudPcWidths(showInUse: true, showUser: true, showServicePlan: false, sideBySide: true, windowWidth: 220);

        Assert.True(narrow.Type <= wide.Type);
        Assert.True(narrow.Name >= 16);
    }

    [Fact]
    public void GetCloudPcWidths_PlentyOfRoom_UsesPreferredMaxSizes()
    {
        var widths = W365CliApp.GetCloudPcWidths(showInUse: true, showUser: false, showServicePlan: false, sideBySide: true, windowWidth: 300);

        Assert.Equal(24, widths.Status);
        Assert.Equal(18, widths.Type);
        Assert.Equal(26, widths.InUse);
    }

    [Fact]
    public void GetCloudPcWidths_NotSideBySide_HasMoreRoomThanSideBySideAtSameWindowWidth()
    {
        var minimum = W365CliApp.GetMinimumSideBySideCloudPcTableWidth(showInUse: true, showUser: true, showServicePlan: false);
        var sideBySide = W365CliApp.GetCloudPcWidths(showInUse: true, showUser: true, showServicePlan: false, sideBySide: true, minimum);
        var standalone = W365CliApp.GetCloudPcWidths(showInUse: true, showUser: true, showServicePlan: false, sideBySide: false, minimum);

        Assert.True(standalone.Name >= sideBySide.Name);
    }
}
