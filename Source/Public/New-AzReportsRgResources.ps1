#requires -Modules ImportExcel

function New-AzReportsRgResources {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Storage Accounts
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Storage Accounts
    .EXAMPLE
        PS C:\> New-AzReportsRgResources -Path .\temp\RgResources.xlsx -Force

        Creates a report of the Azure Storage Accounts and if the Path already exists it overwrites it.
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

        # Name of the Resource Group.
        [Parameter(Mandatory)]
        [string]
        $ResourceGroupName,

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

        $subscription = Get-AzSubscription -SubscriptionId (Get-AzContext).Subscription.Id

        $null = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop

        $resources = Get-AzResource -ResourceGroupName $ResourceGroupName

        Write-Information "Resources Count: $( $resources.Count )"

        $report = @()

        for ($i = 0; $i -lt $resources.Count; $i++) {
            Write-Information "Getting Report Info for Resource: $( $i + 1 ) of $( $resources.Count )"

            $report += [PSCustomObject]@{
                'Subscription Id'     = $subscription.Id
                'Subscription Name'   = $subscription.Name
                'Resource Group Name' = $resources[$i].ResourceGroupName
                Name                  = $resources[$i].Name
                Location              = $resources[$i].Location
                Type                  = $resources[$i].Type
                Id                    = $resources[$i].Id
                ChangedTime           = $resources[$i].ChangedTime
                CreatedTime           = $resources[$i].CreatedTime
                ExtensionResourceName = $resources[$i].ExtensionResourceName
                ExtensionResourceType = $resources[$i].ExtensionResourceType
                Kind                  = $resources[$i].Kind
                ManagedBy             = $resources[$i].ManagedBy
                ParentResource        = $resources[$i].ParentResource
                Plan                  = $resources[$i].Plan
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = $ResourceGroupName
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $report |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
