using Xunit;

namespace W365Cli.Tests;

/// <summary>
/// Covers the Windows App vs. web browser classification backing both the Client Connection Type
/// Report (Reports menu) and the Cloud PC details screen's "Connected via" line -- real field
/// values (ClientFriendlyName, PlatformName) confirmed via captured browser network traffic
/// against a live tenant.
/// </summary>
public class ClientConnectionTypeTests
{
    [Theory]
    [InlineData("Windows App (MSRDC, Store)")]
    [InlineData("Windows App (MSRDC, PerMachine)")]
    public void ClassifyClientConnectionKind_NativeWindowsApp_ReturnsWindowsApp(string clientFriendlyName)
    {
        Assert.Equal("Windows App", W365CliApp.ClassifyClientConnectionKind(clientFriendlyName));
    }

    [Theory]
    [InlineData("Windows App (Web AME)")]
    [InlineData("Browser (Web)")]
    public void ClassifyClientConnectionKind_WebClient_ReturnsWebBrowser(string clientFriendlyName)
    {
        Assert.Equal("Web browser", W365CliApp.ClassifyClientConnectionKind(clientFriendlyName));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("-")]
    public void ClassifyClientConnectionKind_NoData_ReturnsUnknown(string? clientFriendlyName)
    {
        Assert.Equal("Unknown", W365CliApp.ClassifyClientConnectionKind(clientFriendlyName));
    }

    [Fact]
    public void FormatClientConnectionPlain_NativeApp_IgnoresPlatformName()
    {
        // Confirmed on a live tenant: PlatformName comes back as the literal placeholder "<>" for
        // native Windows App sessions -- it must never be appended for those.
        var result = W365CliApp.FormatClientConnectionPlain("Windows App (MSRDC, Store)", "<>");
        Assert.Equal("Windows App (MSRDC, Store)", result);
    }

    [Fact]
    public void FormatClientConnectionPlain_WebWithBrowser_AppendsBrowserName()
    {
        var result = W365CliApp.FormatClientConnectionPlain("Windows App (Web AME)", "Edge");
        Assert.Equal("Windows App (Web AME) (Edge)", result);
    }

    [Theory]
    [InlineData("<>")]
    [InlineData("-")]
    [InlineData(null)]
    public void FormatClientConnectionPlain_WebWithoutBrowser_OmitsBrowserSuffix(string? platformName)
    {
        var result = W365CliApp.FormatClientConnectionPlain("Windows App (Web AME)", platformName);
        Assert.Equal("Windows App (Web AME)", result);
    }

    [Fact]
    public void FormatClientConnectionPlain_NoData_ReturnsDash()
    {
        Assert.Equal("-", W365CliApp.FormatClientConnectionPlain(null, "Edge"));
    }

    [Fact]
    public void FormatClientConnectionMarkup_NativeApp_UsesGreen()
    {
        var result = W365CliApp.FormatClientConnectionMarkup("Windows App (MSRDC, Store)", "<>");
        Assert.Contains("[green]", result, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatClientConnectionMarkup_WebBrowser_UsesYellow()
    {
        var result = W365CliApp.FormatClientConnectionMarkup("Windows App (Web AME)", "Edge");
        Assert.Contains("[yellow]", result, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatClientConnectionMarkup_Unknown_UsesGrey()
    {
        var result = W365CliApp.FormatClientConnectionMarkup(null, null);
        Assert.Contains("[grey]", result, StringComparison.Ordinal);
    }
}
