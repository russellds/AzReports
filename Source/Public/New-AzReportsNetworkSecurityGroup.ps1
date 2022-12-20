#requires -Modules ImportExcel

function New-AzReportsNetworkSecurityGroup {
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

        $nsgs = Get-AzNetworkSecurityGroup

        $nsgReport = @()
        $securityRulesReport = @()

        foreach ($nsg in $nsgs) {
            $nsgReport += [PSCustomObject]@{
                'Resource Group' = $nsg.ResourceGroupName
                Name             = $nsg.Name
                Location         = $nsg.Location
            }

            foreach ($securityRule in $nsg.SecurityRules) {
                $securityRulesReport += [PSCustomObject]@{
                    'NSG Name'                   = $nsg.Name
                    Name                         = $securityRule.Name
                    Description                  = $securityRule.Description
                    Priority                     = $securityRule.Priority
                    Access                       = $securityRule.Access
                    Direction                    = $securityRule.Direction
                    Protocol                     = $securityRule.Protocol
                    'Destination Port Range'     = $securityRule.DestinationPortRange -join ', '
                    'Destination Address Prefix' = ($securityRule.DestinationAddressPrefix | Sort-Object) -join ', '
                    'Source Port Range'          = $securityRule.SourcePortRange -join ', '
                    'Source Address Prefix'      = ($securityRule.SourceAddressPrefix | Sort-Object) -join ', '
                }
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'NSGs'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $nsgReport |
            Sort-Object -Property 'Resource Group', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Security Rules'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $securityRulesReport |
            Sort-Object -Property 'NSG Name', Priority |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -Width 50 -WrapText
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Left
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Left
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -Width 50 -WrapText

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
