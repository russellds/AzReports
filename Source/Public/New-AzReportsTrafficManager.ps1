#requires -Modules ImportExcel

function New-AzReportsTrafficManager {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Traffic Manager
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Traffic Manager
    .EXAMPLE
        PS C:\> New-AzReportsTrafficManager -Path .\temp\TrafficManager.xlsx -Force

        Creates a report of the Azure Traffic Manager and if the Path already exists it overwrites it.
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

        $query = @'
resources
| where type == "microsoft.network/trafficmanagerprofiles"
| order by subscriptionId asc , resourceGroup asc , name asc
'@


        $queryResults = SearchAzGraph -Query $query

        $tmProfiles = [System.Collections.ArrayList]::new()
        $tmEndpoints = [System.Collections.ArrayList]::new()

        foreach ($queryResult in $queryResults) {
            $null = $tmProfiles.Add([PSCustomObject]@{
                    'Subscription Id'                      = $queryResult.subscriptionId
                    'Resource Group'                       = $queryResult.resourceGroup
                    Name                                   = $queryResult.name
                    Status                                 = $queryResult.properties.profileStatus
                    'Routing Method'                       = $queryResult.properties.TrafficRoutingMethod
                    'Max Return'                           = $queryResult.properties.MaxReturn
                    'Traffic View Enrollent'               = $queryResult.properties.trafficViewEnrollmentStatus
                    FQDN                                   = $queryResult.properties.dnsConfig.fqdn
                    'Relative DNS Name'                    = $queryResult.properties.dnsConfig.relativeName
                    TTL                                    = $queryResult.properties.dnsConfig.TTL
                    'Monitor Status'                       = $queryResult.properties.monitorConfig.profileMonitorStatus
                    'Monitor Protocol'                     = $queryResult.properties.monitorConfig.protocol
                    'Monitor Port'                         = $queryResult.properties.monitorConfig.port
                    'Monitor Path'                         = $queryResult.properties.monitorConfig.path
                    'Monitor Interval in Seconds'          = $queryResult.properties.monitorConfig.intervalInSeconds
                    'Monitor Tolerated Number of Failures' = $queryResult.properties.monitorConfig.toleratedNumberOfFailures
                    'Monitor Timeout in Seconds'           = $queryResult.properties.monitorConfig.timeoutInSeconds
                })

            foreach ($endpoint in $queryResult.properties.endpoints) {
                $null = $tmEndpoints.Add([PSCustomObject]@{
                        'Subscription Id'          = $queryResult.subscriptionId
                        'Resource Group'           = $queryResult.resourceGroup
                        'Profile Name'             = $queryResult.name
                        'Endpoint Name'            = $endpoint.name
                        Status                     = $endpoint.properties.endpointStatus
                        Target                     = $endpoint.properties.target
                        Priority                   = $endpoint.properties.priority
                        Weight                     = $endpoint.properties.weight
                        Location                   = $endpoint.properties.endpointLocation
                        'Min Child Endpoints'      = $endpoint.properties.minChildEndpoints
                        'Min Child Endpoints IPv4' = $endpoint.properties.minChildEndpointsIPv4
                        'Min Child Endpoints IPv6' = $endpoint.properties.minChildEndpointsIPv6

                    })
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'Profiles'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $tmProfiles |
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
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 14 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 15 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 16 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 17 -AutoSize -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Endpoints'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $tmEndpoints |
            Sort-Object -Property 'Subscription Id', 'Resource Group', 'Profile Name', 'Endpoint Name' |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -HorizontalAlignment Center

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
