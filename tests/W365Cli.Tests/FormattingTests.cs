using Xunit;

namespace W365Cli.Tests;

public class FormattingTests
{
    [Theory]
    [InlineData("hello", 10, "hello     ")]
    [InlineData("hello world", 8, "hello...")]
    [InlineData("", 4, "    ")]
    public void Fit_PadsOrTruncatesToWidth(string input, int width, string expected)
    {
        var result = W365CliApp.Fit(input, width);
        Assert.Equal(expected, result);
        Assert.Equal(width, result.Length);
    }

    [Theory]
    [InlineData("provisioned", "darkolivegreen3_1")]
    [InlineData("available", "darkolivegreen3_1")]
    [InlineData("provisionedWithWarnings", "orange1")]
    [InlineData("failed", "indianred1")]
    [InlineData("ingraceperiod", "plum1")]
    [InlineData("someUnknownStatus", "grey")]
    [InlineData(null, "grey")]
    public void StatusMarkup_ColorsMatchStatusCategory(string? status, string expectedColor)
    {
        var markup = W365CliApp.StatusMarkup(status);
        Assert.StartsWith($"[{expectedColor}]", markup, StringComparison.Ordinal);
        Assert.EndsWith("[/]", markup, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ready", "khaki1")]
    [InlineData("published", "darkolivegreen3_1")]
    [InlineData("failed", "indianred1")]
    [InlineData("weird", "grey")]
    public void AppStatusMarkup_ColorsMatchStatusCategory(string status, string expectedColor)
    {
        var markup = W365CliApp.AppStatusMarkup(status, 20);
        Assert.StartsWith($"[{expectedColor}]", markup, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("v1.2.3", "1.2.3")]
    [InlineData("1.2.3-beta", "1.2.3")]
    [InlineData("1.2.3+abcdef", "1.2.3")]
    [InlineData("V0.5.38", "0.5.38")]
    public void ParseVersion_StripsPrefixPrereleaseAndMetadata(string input, string expectedVersion)
    {
        var parsed = W365CliApp.ParseVersion(input);
        Assert.NotNull(parsed);
        Assert.Equal(Version.Parse(expectedVersion), parsed);
    }

    [Fact]
    public void ParseVersion_ReturnsNullForGarbage()
    {
        Assert.Null(W365CliApp.ParseVersion("not-a-version"));
    }
}
