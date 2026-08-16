using System.Text.Json;
using Xunit;

namespace W365Cli.Tests;

public class GraphClientUtilityTests
{
    [Theory]
    [InlineData(1073741824L, 1.0)]
    [InlineData(2147483648L, 2.0)]
    [InlineData(null, null)]
    public void ToGb_ConvertsBytesToRoundedGigabytes(long? bytes, double? expected)
    {
        Assert.Equal(expected, W365GraphClient.ToGb(bytes));
    }

    [Fact]
    public void GetFirst_ReturnsFirstNonBlankMatch()
    {
        var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["DisplayName"] = "",
            ["UPN"] = "user@contoso.com"
        };

        Assert.Equal("user@contoso.com", W365GraphClient.GetFirst(fields, "DisplayName", "UPN"));
    }

    [Fact]
    public void GetFirst_ReturnsNullWhenNoNamesMatch()
    {
        var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Assert.Null(W365GraphClient.GetFirst(fields, "DisplayName", "UPN"));
    }

    [Fact]
    public void JoinSummary_SkipsBlankAndDashValues()
    {
        Assert.Equal("A | B", W365GraphClient.JoinSummary("A", "", null, "-", "B"));
    }

    [Fact]
    public void JoinSummary_AllBlank_ReturnsDash()
    {
        Assert.Equal("-", W365GraphClient.JoinSummary("", null, "-"));
    }

    [Theory]
    [InlineData("true", "Yes")]
    [InlineData("false", "No")]
    [InlineData(null, "Unknown")]
    [InlineData("", "Unknown")]
    [InlineData("-", "Unknown")]
    public void FormatBoolean_MapsToYesNoUnknown(string? input, string expected)
    {
        Assert.Equal(expected, W365GraphClient.FormatBoolean(input));
    }

    [Fact]
    public void JsonToString_HandlesAllValueKinds()
    {
        using var doc = JsonDocument.Parse("""{"s":"hi","n":42,"t":true,"f":false,"nul":null}""");
        var root = doc.RootElement;
        Assert.Equal("hi", W365GraphClient.JsonToString(root.GetProperty("s")));
        Assert.Equal("42", W365GraphClient.JsonToString(root.GetProperty("n")));
        Assert.Equal("true", W365GraphClient.JsonToString(root.GetProperty("t")));
        Assert.Equal("false", W365GraphClient.JsonToString(root.GetProperty("f")));
        Assert.Equal("-", W365GraphClient.JsonToString(root.GetProperty("nul")));
    }
}
