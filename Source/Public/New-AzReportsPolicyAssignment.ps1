#requires -Modules ImportExcel

function New-AzReportsPolicyAssignment {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Policy Assignment
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Policy Assignment
    .EXAMPLE
        PS C:\> New-AzReportsPolicyAssignment -Path .\temp\SecurityCenterBuiltIn.xlsx -Name SecurityCenterBuiltIn -Force

        Creates a report of the Azure Policy Assignment and if the Path already exists it overwrites it.
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

        [string]
        $Name,

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

        if ($Name) {
            $policyAssignments = Get-AzPolicyAssignment -Name $Name
        } else {
            $policyAssignments = Get-AzPolicyAssignment
        }

        $objects = @()

        foreach ( $policyAssignment in $policyAssignments) {
            if ($policyAssignment.Properties.Parameters) {
                $objects += GetPolicyAssignmentParameters -PolicyAssignment $policyAssignment
            } else {
                Write-Information "Policy Assignment: $( $policyAssignment.Properties.DisplayName ) - has no parameters."

                $objects += [PSCustomObject]@{
                    Name             = $policyAssignment.Name
                    'Display Name'   = $policyAssignment.Properties.DisplayName
                    Scope            = $policyAssignment.Properties.Scope
                    'Parameter Name' = $null
                    Value            = $null
                }
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'PolicyAssignment'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $objects |
            Sort-Object -Property Category, Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
