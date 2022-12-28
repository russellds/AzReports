#requires -Modules ImportExcel

function New-AzReportsAppInsights {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure ApplicationInsights
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure ApplicationInsights
    .EXAMPLE
        PS C:\> New-AzReportsAppInsights -Path .\temp\AppInsights.xlsx -Force

        Creates a report of the Azure ApplicationInsights and if the Path already exists it overwrites it.
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

        $currentSubscription = Get-AzSubscription -SubscriptionId (Get-AzContext).Subscription.Id

        if ($Current) {
            $subscriptions = $currentSubscription
        } else {
            $subscriptions = Get-AzSubscription
        }

        $appInsightsReport = @()

        foreach ($subscription in $subscriptions) {

            Write-Information "Setting Azure Context to Subscription: $( $subscription.Name )"
            $null = Set-AzContext -SubscriptionId $subscription.Id

            $appInsights = Get-AzApplicationInsights

            Write-Information "AppInsights Count: $( $appInsights.Count )"

            foreach ($appInsight in $appInsights) {
                $insight = Get-AzApplicationInsights -ResourceGroupName $appInsight.ResourceGroupName -Name $appInsight.Name
                $appInsightsReport += [PSCustomObject]@{
                    'Subscription Id'                     = $subscription.Id
                    'Subscription Name'                   = $subscription.Name
                    'Resource Group Name'                 = $insight.ResourceGroupName
                    Name                                  = $insight.Name
                    Id                                    = $insight.Id
                    'App Id'                              = $insight.AppId
                    'Application Id'                      = $insight.ApplicationId
                    'Application Type'                    = $insight.ApplicationType
                    'Creation Date'                       = $insight.CreationDate
                    'Disable IP Masking'                  = $insight.DisableIPMasking
                    'Disable Local Auth'                  = $insight.DisableLocalAuth
                    'Immediate Purge Data on 30 Day'      = $insight.ImmediatePurgeDataOn30Day
                    'Ingestion Mode'                      = $insight.IngestionMode
                    Kind                                  = $insight.Kind
                    Location                              = $insight.Location
                    'Private Link Scoped Resource'        = $insight.PrivateLinkScopedResource
                    'Provisioning State'                  = $insight.ProvisioningState
                    'Public Network Access for Ingestion' = $insight.PublicNetworkAccessForIngestion
                    'Public Network Access for Query'     = $insight.PublicNetworkAccessForQuery
                    'Retention In Day'                    = $insight.RetentionInDay
                    'Sampling Percentage'                 = $insight.SamplingPercentage
                    Type                                  = $insight.Type
                    'Workspace Resource Id'               = $insight.WorkspaceResourceId
                }
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'AppInsights'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $appInsightsReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
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

        if ((Get-AzContext).Subscription.Id -ne $currentSubscription.Id) {
            Write-Information "Setting Azure Context to Subscription: $( $currentSubscription.Name )"
            $null = Set-AzContext -SubscriptionId $currentSubscription.Id
        }
    } catch {
        throw $PSItem
    }
}
