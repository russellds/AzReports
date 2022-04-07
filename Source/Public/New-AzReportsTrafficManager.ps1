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

        $tmProfiles = Get-AzTrafficManagerProfile

        $customTmProfileProperties = @(
            'ResourceGroupName',
            'Name',
            'RelativeDnsName',
            'Ttl',
            'ProfileStatus',
            'TrafficRoutingMethod',
            'MonitorProtocol',
            'MonitorPort',
            'MonitorPath',
            'MonitorIntervalInSeconds',
            'MonitorTimeoutInSeconds',
            'MonitorToleratedNumberOfFailures',
            'MaxReturn'
        )

        $customTmProfileObjects = $tmProfiles |
            Select-Object -Property $customTmProfileProperties

        $customEndpointProperties = @(
            'ResourceGroupName',
            'ProfileName',
            'Name',
            'Type',
            'Target',
            'EndpointStatus',
            'Weight',
            'Priority',
            'Location',
            'EndpointMonitorStatus',
            'MinChildEndpoints',
            'MinChildEndpointsIPv4',
            'MinChildEndpointsIPv6',
            'GeoMapping',
            'SubnetMapping',
            'CustomHeaders'
        )

        $customEndpointObjects = $tmProfiles.Endpoints |
            Select-Object -Property $customEndpointProperties

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'Profiles'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $customTmProfileObjects |
            Sort-Object -Property ResourecGroupName, Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Endpoints'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $customEndpointObjects |
            Sort-Object -Property ResourecGroupName, ProfileName, Name |
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
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 14 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 15 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 16 -AutoSize

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
