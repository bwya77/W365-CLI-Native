namespace W365Cli.Tests;

public class FrameWriterTests
{
    private const string Hi = "\x1b[30;48;2;88;166;255m";
    private const string Off = "\x1b[0m";

    private static string[] Lines(string output) =>
        output.Replace("\x1b[H", "").Replace("\x1b[K", "").Replace("\x1b[J", "")
            .Split('\n');

    private static string Frame(int count, int selected, bool trailing = true)
    {
        var rows = Enumerable.Range(0, count).Select(i =>
            i == 0 ? $"{Hi}nav{Off}" : i == selected ? $"{Hi}row {i}{Off}" : $"row {i}");
        return string.Join("\n", rows) + (trailing ? "\n" : "");
    }

    [Fact]
    public void FittingFrameIsKeptIntactWithoutClearing()
    {
        var offset = 0;
        var output = FrameWriter.Compose(Frame(5, 2), 24, ref offset);
        Assert.StartsWith("\x1b[H", output);
        Assert.EndsWith("\x1b[J", output);
        Assert.DoesNotContain("\x1b[2J", output);
        Assert.Equal(6, Lines(output).Length); // 5 rows plus the trailing newline's empty entry
        Assert.Contains("row 4", output);
    }

    [Theory]
    [InlineData(10, 1)]
    [InlineData(10, 20)]
    [InlineData(10, 39)]
    [InlineData(6, 30)]
    public void TallFrameIsClippedToWindowKeepingNavAndSelection(int height, int selected)
    {
        var offset = 0;
        var output = FrameWriter.Compose(Frame(40, selected), height, ref offset);
        var lines = Lines(output);
        Assert.Equal(height, lines.Length);
        Assert.Contains("nav", lines[0]);
        Assert.Contains(lines, line => line.Contains($"row {selected}"));
    }

    [Fact]
    public void ScrollOffsetIsStickyWhileSelectionStaysVisible()
    {
        var offset = 0;
        FrameWriter.Compose(Frame(40, 25), 10, ref offset);
        var before = offset;
        FrameWriter.Compose(Frame(40, 24), 10, ref offset);
        Assert.Equal(before, offset);
    }
}


