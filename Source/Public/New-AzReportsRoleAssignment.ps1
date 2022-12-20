#requires -Modules ImportExcel

function New-AzReportsRoleAssignment {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Role Assignments
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Role Assignments
    .EXAMPLE
        PS C:\> New-AzReportsRoleAssignment -Path .\temp\RoleAssignment.xlsx -Force

        Creates a report of the Azure Role Assignments and if the Path already exists it overwrites it.
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

        # Only generate report for the current Azure Subscription.
        [switch]
        $Current,

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

        if ($Current) {
            $subscriptions = Get-AzSubscription -SubscriptionId (Get-AzContext).Subscription.Id
        } else {
            $subscriptions = Get-AzSubscription
        }

        $rolesReport = @()

        foreach ($subscription in $subscriptions) {

            Write-Information "Setting Azure Context to Subscription: $( $subscription.Name )" -InformationAction Continue
            $context = Set-AzContext -SubscriptionId $subscription.Id

            $assignments = Get-AzRoleAssignment

            foreach ($assignment in $assignments) {
                $customRole = (Get-AzRoleDefinition -Name $assignment.RoleAssignmentName).IsCustom

                $rolesReport += [PSCustomObject]@{
                    'Subscription Id'      = $subscription.Id
                    'Subscription Name'    = $subscription.Name
                    'Display Name'         = $assignment.DisplayName
                    'Sign-In Name'         = $assignment.SignInName
                    'Object Type'          = $assignment.ObjectType
                    'Role Definition Name' = $assignment.RoleDefinitionName
                    'Custom Role'          = $customRole
                    'Scope'                = $assignment.Scope
                }
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'RoleAssignments'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $rolesReport |
            Sort-Object -Property 'Subscription Name', 'Display Name', 'Role Definition Name' |
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

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
