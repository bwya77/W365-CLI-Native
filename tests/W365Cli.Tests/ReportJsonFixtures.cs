namespace W365Cli.Tests;

/// <summary>
/// Real (redacted) Microsoft Graph report payload shapes captured from a live tenant during
/// development, reused here so the report-parsing tests are grounded in actual Graph response
/// shapes instead of invented ones.
/// </summary>
internal static class ReportJsonFixtures
{
    /// <summary>
    /// Shape of the columnar {"Schema": [...], "Values": [[...], ...]} response returned by the
    /// undocumented deviceManagement/virtualEndpoint/reports/retrieveCloudPcTroubleshootReports
    /// (troubleshootConnectionConfigurationOfViewDataTableV1Report) and similar reporting
    /// endpoints. Two rows, redacted UPNs/IDs.
    /// </summary>
    public const string ConnectionHistoryReport = """
    {
      "Schema": [
        { "Column": "ActivityId" },
        { "Column": "SessionBeginTime" },
        { "Column": "SessionEndTime" },
        { "Column": "UPN" },
        { "Column": "ManagedDeviceName" },
        { "Column": "CloudPCId" },
        { "Column": "ClientOS" },
        { "Column": "TransportType" }
      ],
      "Values": [
        [
          "activity-1",
          "2024-06-01T13:00:00Z",
          "2024-06-01T14:30:00Z",
          "user1@contoso.com",
          "CPC-USER1",
          "11111111-1111-1111-1111-111111111111",
          "Windows 11",
          "RDP Shortpath"
        ],
        [
          "activity-2",
          "2024-06-02T09:15:00Z",
          "2024-06-02T09:45:00Z",
          "user2@contoso.com",
          "CPC-USER2",
          "22222222-2222-2222-2222-222222222222",
          "macOS",
          "TCP"
        ]
      ]
    }
    """;

    /// <summary>
    /// Real shape from the frontlineLicenseHourlyUsageReport capture: its useful columns
    /// (LicenseCount, ClaimedLicenseCount, SkuLicenseCount) are not part of the fixed
    /// connection-history summary field set, which is exactly the scenario
    /// ParseReportRowsAdaptive was built to handle generically.
    /// </summary>
    public const string FrontlineLicenseHourlyUsageReport = """
    {
      "Schema": [
        { "Column": "DisplayName" },
        { "Column": "Timestamp" },
        { "Column": "LicenseCount" },
        { "Column": "ClaimedLicenseCount" },
        { "Column": "SkuLicenseCount" },
        { "Column": "UniqueId" },
        { "Column": "IngestedTimestamp" }
      ],
      "Values": [
        [
          "Frontline Flex Pool",
          "2024-06-01T00:00:00Z",
          "10",
          "7",
          "10",
          "33333333-3333-3333-3333-333333333333",
          "2024-06-01T00:05:00Z"
        ],
        [
          "Frontline Flex Pool",
          "2024-06-01T01:00:00Z",
          "10",
          "9",
          "10",
          "44444444-4444-4444-4444-444444444444",
          "2024-06-01T01:05:00Z"
        ]
      ]
    }
    """;

    public const string EmptyReport = """
    {
      "Schema": [],
      "Values": []
    }
    """;

    public const string MalformedReport = """
    {
      "SomethingElse": true
    }
    """;
}
