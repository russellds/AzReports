#requires -Modules ImportExcel, Az.ResourceGraph

function New-AzReportsSqlDatabase {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Sql Database
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Sql Database
    .EXAMPLE
        PS C:\> New-AzReportsSqlDatabase -Path .\temp\SqlDatabase.xlsx -Force

        Creates a report of the Azure Sql Database and if the Path already exists it overwrites it.
    .INPUTS
        None
    .OUTPUTS
        Excel Spreadsheet
    #>
    [CmdletBinding()]
    param(
        # Path to create the Excel report. Must end with '.xlsx'.
        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        # Do not automatically open the generated Excel spreadsheet.
        [switch]
        $NoInvoke,

        # Overwrite existing Excel spreadsheet.
        [switch]
        $Force
    )
    $InformationPreference = 'Continue'
    $env:SuppressAzurePowerShellBreakingChangeWarnings = 'true'

    try {
        CheckAzContext

        CheckPath -Path $Path -Extension '.xlsx' -Force:$Force -ErrorAction Stop

        $serverQuery = @'
resources
| where type == "microsoft.sql/servers"
| join kind=inner (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project subscriptionId, subscriptionName = name)
    on subscriptionId
| project-away subscriptionId1
'@


        $serverQueryResults = SearchAzGraph -Query $serverQuery

        $sqlServers = [System.Collections.ArrayList]::new()

        foreach ($serverQueryResult in $serverQueryResults) {
            $null = $sqlServers.Add([PSCustomObject]@{
                    'Subscription Id'                  = $serverQueryResult.subscriptionId
                    'Subscription Name'                = $serverQueryResult.subscriptionName
                    'Resource Group'                   = $serverQueryResult.resourceGroup
                    Name                               = $serverQueryResult.name
                    Location                           = $serverQueryResult.location
                    'Administrator Login'              = $serverQueryResult.properties.administratorLogin
                    'AAD Administrator'                = $serverQueryResult.properties.administrators.login
                    'FQDN'                             = $serverQueryResult.properties.fullyQualifiedDomainName
                    Kind                               = $serverQueryResult.kind
                    'Public Network Access'            = $serverQueryResult.properties.publicNetworkAccess
                    #'Private Endpoint Connections'     = $serverQueryResult.privateEndpointConnections.ToString()
                    'Restrict Outbound Network Access' = $serverQueryResult.properties.restrictOutboundNetworkAccess
                    State                              = $serverQueryResult.properties.state
                    Version                            = $serverQueryResult.properties.version
                })
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'SQL Servers'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $sqlServers |
            Sort-Object -Property 'Subscription Id', 'Resource Group', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 14 -AutoSize

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
