#requires -Modules ImportExcel

function New-AzReportsStorageAccount {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Storage Accounts
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Storage Accounts
    .EXAMPLE
        PS C:\> New-AzReportsStorageAccount -Path .\temp\StorageAccount.xlsx -Force

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

        $storageAccountReport = @()

        foreach ($subscription in $subscriptions) {

            Write-Information "Setting Azure Context to Subscription: $( $subscription.Name )"
            $null = Set-AzContext -SubscriptionId $subscription.Id

            $storageAccounts = Get-AzStorageAccount

            Write-Information "Storage Account Count: $( $storageAccounts.Count )"

            foreach ($storageAccount in $storageAccounts) {
                $storageAccountReport += [PSCustomObject]@{
                    'Subscription Id'                 = $subscription.Id
                    'Subscription Name'               = $subscription.Name
                    'Resource Group Name'             = $storageAccount.ResourceGroupName
                    Name                              = $storageAccount.StorageAccountName
                    Location                          = $storageAccount.Location
                    Kind                              = $storageAccount.Kind
                    Sku                               = $storageAccount.Sku.Name
                    AccessTier                        = $storageAccount.AccessTier
                    AllowBlobPublicAccess             = $storageAccount.AllowBlobPublicAccess
                    AllowCrossTenantReplication       = $storageAccount.AllowCrossTenantReplication
                    AllowSharedKeyAccess              = $storageAccount.AllowSharedKeyAccess
                    AzureFilesIdentityBasedAuth       = $storageAccount.AzureFilesIdentityBasedAuth
                    BlobRestoreStatus                 = $storageAccount.BlobRestoreStatus
                    CreationTime                      = $storageAccount.CreationTime
                    CustomDomain                      = $storageAccount.CustomDomain
                    EnableHierarchicalNamespace       = $storageAccount.EnableHierarchicalNamespace
                    EnableHttpsTrafficOnly            = $storageAccount.EnableHttpsTrafficOnly
                    EnableLocalUser                   = $storageAccount.EnableLocalUser
                    EnableNfsV3                       = $storageAccount.EnableNfsV3
                    EnableSftp                        = $storageAccount.EnableSftp
                    FailoverInProgress                = $storageAccount.FailoverInProgress
                    GeoReplicationStats               = $storageAccount.GeoReplicationStats
                    Identity                          = $storageAccount.Identity
                    ImmutableStorageWithVersioning    = $storageAccount.ImmutableStorageWithVersioning
                    Key1CreationTime                  = $storageAccount.KeyCreationTime.Key1
                    Key2CreationTime                  = $storageAccount.KeyCreationTime.Key2
                    KeyPolicy                         = $storageAccount.KeyPolicy
                    LargeFileSharesState              = $storageAccount.LargeFileSharesState
                    LastGeoFailoverTime               = $storageAccount.LastGeoFailoverTime
                    MinimumTlsVersion                 = $storageAccount.MinimumTlsVersion
                    PrimaryLocation                   = $storageAccount.PrimaryLocation
                    ProvisioningState                 = $storageAccount.ProvisioningState
                    PublicNetworkAccess               = $storageAccount.PublicNetworkAccess
                    RoutingPreference                 = $storageAccount.RoutingPreference
                    SasPolicy                         = $storageAccount.SasPolicy
                    SecondaryLocation                 = $storageAccount.SecondaryLocation
                    StatusOfPrimary                   = $storageAccount.StatusOfPrimary
                    StatusOfSecondary                 = $storageAccount.StatusOfSecondary
                    StorageAccountSkuConversionStatus = $storageAccount.StorageAccountSkuConversionStatus
                    ResourceId                        = $storageAccount.Id
                }
            }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'StorageAccount'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $storageAccountReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

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
