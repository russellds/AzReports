#requires -Modules ImportExcel

function New-AzReportsAppServices {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure App Services
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure App Services
    .EXAMPLE
        PS C:\> New-AzReportsAppServices -Path .\temp\AppServices.xlsx -Force

        Creates a report of the Azure App Services and if the Path already exists it overwrites it.
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

        $report = @()

        foreach ($subscription in $subscriptions) {

            Write-Information "Setting Azure Context to Subscription: $( $subscription.Name )"
            $null = Set-AzContext -SubscriptionId $subscription.Id

            $plans = Get-AzAppServicePlan

            Write-Information "App Service Plans Count: $( $plans.Count )"

            foreach ($plan in $plans) {
                $appServices = Get-AzWebApp -AppServicePlan $plan

                Write-Information "App Service Plan - $( $plan.Name ) - App Services Count: $( $appServices.Count )"

                foreach ($appService in $appServices) {
                    $service = Get-AzWebApp -ResourceGroupName $appService.ResourceGroup -Name $appService.Name

                    Write-Information "Processing App Service - $( $service.Name )..."

                    if ($service.VirtualNetworkSubnetId) {
                        $vnetName = $service.VirtualNetworkSubnetId.Split('/')[8]
                        $subnetName = $service.VirtualNetworkSubnetId.Split('/')[-1]
                    } else {
                        $vnetName = $null
                        $subnetName = $null
                    }

                    $report += [PSCustomObject]@{
                        'Subscription Id'                              = $subscription.Id
                        'Subscription Name'                            = $subscription.Name
                        'Plan Resource Group Name'                     = $plan.ResourceGroup
                        'Plan Name'                                    = $plan.Name
                        'Plan Kind'                                    = $plan.Kind
                        'Plan Sku'                                     = $plan.Sku.Name
                        'Resource Group Name'                          = $service.ResourceGroup
                        Name                                           = $service.Name
                        Kind                                           = $service.Kind
                        Location                                       = $service.Location
                        'Default Host Name'                            = $service.DefaultHostName
                        'Https Only'                                   = $service.HttpsOnly
                        'Virtual Network Name'                         = $vnetName
                        'Subnet Name'                                  = $subnetName
                        'Site Config AlwaysOn'                         = $service.SiteConfig.AlwaysOn
                        'Site Config FtpsState'                        = $service.SiteConfig.FtpsState
                        'Site Config Http20Enabled'                    = $service.SiteConfig.Http20Enabled
                        'Site Config HttpLoggingEnabled'               = $service.SiteConfig.HttpLoggingEnabled
                        'Site Config IpSecurityRestrictions'           = $service.SiteConfig.IpSecurityRestrictions.Name -join ', '
                        'Site Config JavaVersion'                      = $service.SiteConfig.JavaVersion
                        'Site Config LinuxFxVersion'                   = $service.SiteConfig.LinuxFxVersion
                        'Site Config MinTlsVersion'                    = $service.SiteConfig.MinTlsVersion
                        'Site Config NetFrameworkVersion'              = $service.SiteConfig.NetFrameworkVersion
                        'Site Config NodeVersion'                      = $service.SiteConfig.NodeVersion
                        'Site Config PhpVersion'                       = $service.SiteConfig.PhpVersion
                        'Site Config PowerShellVersion'                = $service.SiteConfig.PowerShellVersion
                        'Site Config PythonVersion'                    = $service.SiteConfig.PythonVersion
                        'Site Config RemoteDebuggingEnabled'           = $service.SiteConfig.RemoteDebuggingEnabled
                        'Site Config RequestTracingEnabled'            = $service.SiteConfig.RequestTracingEnabled
                        'Site Config ScmIpSecurityRestrictions'        = $service.SiteConfig.ScmIpSecurityRestrictions.Name -join ', '
                        'Site Config ScmIpSecurityRestrictionsUseMain' = $service.SiteConfig.ScmIpSecurityRestrictionsUseMain
                        'Site Config ScmMinTlsVersion'                 = $service.SiteConfig.ScmMinTlsVersion
                        'Site Config ScmType'                          = $service.SiteConfig.ScmType
                        'Site Config Use32BitWorkerProcess'            = $service.SiteConfig.Use32BitWorkerProcess
                        'Site Config VnetRouteAllEnabled'              = $service.SiteConfig.VnetRouteAllEnabled
                        'Site Config WebSocketsEnabled'                = $service.SiteConfig.WebSocketsEnabled
                        'Site Config WindowsFxVersion'                 = $service.SiteConfig.WindowsFxVersion
                    }
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

        $excel = $report |
            Sort-Object -Property 'Subscription Name', 'App Service Plan Resource Group Name', 'App Service Plan Name', 'Resource Group Name', Name |
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
        Set-ExcelColumn -Worksheet $workSheet -Column 15 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 16 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 17 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 18 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 19 -HorizontalAlignment Left -Width 50 -WrapText
        Set-ExcelColumn -Worksheet $workSheet -Column 20 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 21 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 22 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 23 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 24 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 25 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 26 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 27 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 28 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 29 -HorizontalAlignment Left -Width 50 -WrapText
        Set-ExcelColumn -Worksheet $workSheet -Column 30 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 31 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 32 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 33 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 34 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 35 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 36 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 37 -AutoSize

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
