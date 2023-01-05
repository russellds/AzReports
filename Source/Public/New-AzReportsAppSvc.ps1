#requires -Modules ImportExcel

function New-AzReportsAppSvc {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure App Service Plans and App Services.
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure App Service Plans and App Services.
    .EXAMPLE
        PS C:\> New-AzReportsAppAppSvc -Path .\temp\AppAppSvc.xlsx -Force

        Creates a report of the Azure App Service Plans and App Services and if the Path already exists it overwrites it.
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

        $appSvcPlanReport = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
        $webAppReport = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
        $webAppHostSslStatesReport = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
        $webAppSiteConfigReport = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
        $webAppSiteConfigAppSettingsReport = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()

        foreach ($subscription in $subscriptions) {

            Write-Information "Setting Azure Context to Subscription: $( $subscription.Name )"
            $null = Set-AzContext -SubscriptionId $subscription.Id

            $appSvcPlans = Get-AzAppServicePlan

            Write-Information "AppSvc Plan Count: $( $appSvcPlans.Count )"

            $appSvcPlans |
                ForEach-Object -ThrottleLimit 5 -Parallel {
                    $localSubscription = $using:subscription
                    $localAppSvcPlanReport = $using:appSvcPlanReport
                    $localWebAppReport = $using:webAppReport
                    $localWebAppHostSslStatesReport = $using:webAppHostSslStatesReport
                    $localWebAppSiteConfigReport = $using:webAppSiteConfigReport

                    $object = [PSCustomObject]@{
                        'Subscription Id'              = $localSubscription.Id
                        'Subscription Name'            = $localSubscription.Name
                        'Resource Group Name'          = $_.ResourceGroup
                        Name                           = $_.Name
                        Location                       = $_.Location
                        Sku                            = '{0} ({1})' -f $_.Sku.Tier, $_.Sku.Size
                        'Admin Site Name'              = $_.AdminSiteName
                        'ElasticScale Enabled'         = $_.ElasticScaleEnabled
                        'Extended Location'            = $_.ExtendedLocation
                        'Free Offer Expiration Time'   = $_.FreeOfferExpirationTime
                        'Geo Region'                   = $_.GeoRegion
                        'Hosting Environment Profile'  = $_.HostingEnvironmentProfile
                        'HyperV'                       = $_.HyperV
                        'Is Spot'                      = $_.IsSpot
                        'Is Xenon'                     = $_.IsXenon
                        Kind                           = $_.Kind
                        'Kube Environment Profile'     = $_.KubeEnvironmentProfile
                        'Maximum Elastic Worker Count' = $_.MaximumElasticWorkerCount
                        'MaximumNumberOfWorkers'       = $_.MaximumNumberOfWorkers
                        'Number Of Sites'              = $_.NumberOfSites
                        'Per Site Scaling'             = $_.PerSiteScaling
                        'Provisioning State'           = $_.ProvisioningState
                        'Reserved'                     = $_.Reserved
                        'Spot Expiration Time'         = $_.SpotExpirationTime
                        'Status'                       = $_.Status.ToString()
                        'Target Worker Count'          = $_.TargetWorkerCount
                        'Target Worker Size Id'        = $_.TargetWorkerSizeId
                        Type                           = $_.Type
                        'Worker Tier Name'             = $_.WorkerTierName
                        'Resource Id'                  = $_.Id
                    }

                    $localAppSvcPlanReport.Add($object)

                    $webApps = Get-AzWebApp -AppServicePlan $_

                    Write-Information "Web Apps Count: $( $webApps.Count )"

                    foreach ($webApp in $webApps) {

                        $app = Get-AzWebApp -ResourceGroupName $webApp.ResourceGroup -Name $webApp.Name

                        $object = [PSCustomObject]@{
                            'Subscription Id'             = $localSubscription.Id
                            'Subscription Name'           = $localSubscription.Name
                            'Resource Group Name'         = $app.ResourceGroup
                            Name                          = $app.Name
                            Location                      = $app.Location
                            State                         = $app.State
                            'AvailabilityState'           = $app.AvailabilityState
                            'AzureStorageAccounts'        = $app.AzureStorageAccounts
                            'AzureStoragePath'            = $app.AzureStoragePath
                            'ClientAffinityEnabled'       = $app.ClientAffinityEnabled
                            'ClientCertEnabled'           = $app.ClientCertEnabled
                            'ClientCertExclusionPaths'    = $app.ClientCertExclusionPaths
                            'ClientCertMode'              = $app.ClientCertMode
                            'CloningInfo'                 = $app.CloningInfo
                            'ContainerSize'               = $app.ContainerSize
                            'CustomDomainVerificationId'  = $app.CustomDomainVerificationId
                            'DailyMemoryTimeQuota'        = $app.DailyMemoryTimeQuota
                            'DefaultHostName'             = $app.DefaultHostName
                            'Enabled'                     = $app.Enabled
                            'EnabledHostNames'            = $app.EnabledHostNames -join ', '
                            'ExtendedLocation'            = $app.ExtendedLocation
                            'GitRemoteName'               = $app.GitRemoteName
                            'GitRemoteUri'                = $app.GitRemoteUri
                            'GitRemoteUsername'           = $app.GitRemoteUsername
                            'HostingEnvironmentProfile'   = $app.HostingEnvironmentProfile
                            'HostNames'                   = $app.HostNames -join ', '
                            'HostNamesDisabled'           = $app.HostNamesDisabled
                            'HttpsOnly'                   = $app.HttpsOnly
                            'HyperV'                      = $app.HyperV
                            'Identity'                    = $app.Identity
                            'InProgressOperationId'       = $app.InProgressOperationId
                            'IsDefaultContainer'          = $app.IsDefaultContainer
                            'IsXenon'                     = $app.IsXenon
                            'KeyVaultReferenceIdentity'   = $app.KeyVaultReferenceIdentity
                            'Kind'                        = $app.Kind
                            'LastModifiedTimeUtc'         = $app.LastModifiedTimeUtc
                            'MaxNumberOfWorkers'          = $app.MaxNumberOfWorkers
                            'OutboundIpAddresses'         = $app.OutboundIpAddresses
                            'PossibleOutboundIpAddresses' = $app.PossibleOutboundIpAddresses
                            'RedundancyMode'              = $app.RedundancyMode
                            'RepositorySiteName'          = $app.RepositorySiteName
                            'Reserved'                    = $app.Reserved
                            'ScmSiteAlsoStopped'          = $app.ScmSiteAlsoStopped
                            'ServerFarmId'                = $app.ServerFarmId
                            'SlotSwapStatus'              = $app.SlotSwapStatus
                            'StorageAccountRequired'      = $app.StorageAccountRequiredss
                            'SuspendedTill'               = $app.SuspendedTill
                            'TargetSwapSlot'              = $app.TargetSwapSlot
                            'TrafficManagerHostNames'     = $app.TrafficManagerHostNames -join ', '
                            'Type'                        = $app.Type
                            'UsageState'                  = $app.UsageState
                            'VirtualNetworkSubnetId'      = $app.VirtualNetworkSubnetId
                            'Resource Id'                 = $app.Id
                        }

                        $localWebAppReport.Add($object)

                        foreach ($hostSslState in $app.HostNameSslStates) {
                            $object = [PSCustomObject]@{
                                'Subscription Id'     = $localSubscription.Id
                                'Subscription Name'   = $localSubscription.Name
                                'Resource Group Name' = $app.ResourceGroup
                                Name                  = $app.Name
                                'Host Type'           = $hostSslState.HostType
                                URI                   = $hostSslState.Name
                                'Ssl State'           = $hostSslState.SslState
                                Thumbprint            = $hostSslState.Thumbprint
                                'To Update'           = $hostSslState.ToUpdate
                                'Virtual IP'          = $hostSslState.VirtualIP
                            }

                            $localWebAppHostSslStatesReport.Add($object)
                        }

                        $webAppSiteConfigObject = [PSCustomObject]@{
                            'Subscription Id'                        = $localSubscription.Id
                            'Subscription Name'                      = $localSubscription.Name
                            'Resource Group Name'                    = $app.ResourceGroup
                            Name                                     = $app.Name
                            'AcrUseManagedIdentityCreds'             = $app.SiteConfig.AcrUseManagedIdentityCreds
                            'AcrUserManagedIdentityID'               = $app.SiteConfig.AcrUserManagedIdentityID
                            'AlwaysOn'                               = $app.SiteConfig.AlwaysOn
                            'ApiDefinition'                          = $app.SiteConfig.ApiDefinition
                            'ApiManagementConfig'                    = $app.SiteConfig.ApiManagementConfig
                            'AppCommandLine'                         = $app.SiteConfig.AppCommandLine
                            'AutoHealEnabled'                        = $app.SiteConfig.AutoHealEnabled
                            'AutoHealRules'                          = $app.SiteConfig.AutoHealRules
                            'AutoSwapSlotName'                       = $app.SiteConfig.AutoSwapSlotName
                            'AzureStorageAccounts'                   = $app.SiteConfig.AzureStorageAccounts
                            'ConnectionStrings'                      = $app.SiteConfig.ConnectionStrings
                            'Cors'                                   = $app.SiteConfig.Cors
                            'DefaultDocuments'                       = $app.SiteConfig.DefaultDocuments
                            'DetailedErrorLoggingEnabled'            = $app.SiteConfig.DetailedErrorLoggingEnabled
                            'DocumentRoot'                           = $app.SiteConfig.DocumentRoot
                            'Experiments'                            = $app.SiteConfig.Experiments
                            'FtpsState'                              = $app.SiteConfig.FtpsState
                            'FunctionAppScaleLimit'                  = $app.SiteConfig.FunctionAppScaleLimit
                            'FunctionsRuntimeScaleMonitoringEnabled' = $app.SiteConfig.FunctionsRuntimeScaleMonitoringEnabled
                            'HandlerMappings'                        = $app.SiteConfig.HandlerMappings
                            'HealthCheckPath'                        = $app.SiteConfig.HealthCheckPath
                            'Http20Enabled'                          = $app.SiteConfig.Http20Enabled
                            'HttpLoggingEnabled'                     = $app.SiteConfig.HttpLoggingEnabled
                            'IpSecurityRestrictions'                 = $app.SiteConfig.IpSecurityRestrictions
                            'JavaContainer'                          = $app.SiteConfig.JavaContainer
                            'JavaContainerVersion'                   = $app.SiteConfig.JavaContainerVersion
                            'JavaVersion'                            = $app.SiteConfig.JavaVersion
                            'KeyVaultReferenceIdentity'              = $app.SiteConfig.KeyVaultReferenceIdentity
                            'Limits'                                 = $app.SiteConfig.Limits
                            'LinuxFxVersion'                         = $app.SiteConfig.LinuxFxVersion
                            'LoadBalancing'                          = $app.SiteConfig.LoadBalancing
                            'LocalMySqlEnabled'                      = $app.SiteConfig.LocalMySqlEnabled
                            'LogsDirectorySizeLimit'                 = $app.SiteConfig.LogsDirectorySizeLimit
                            'MachineKey'                             = $app.SiteConfig.MachineKey
                            'ManagedPipelineMode'                    = $app.SiteConfig.ManagedPipelineMode
                            'ManagedServiceIdentityId'               = $app.SiteConfig.ManagedServiceIdentityId
                            'MinimumElasticInstanceCount'            = $app.SiteConfig.MinimumElasticInstanceCount
                            'MinTlsVersion'                          = $app.SiteConfig.MinTlsVersion
                            'NetFrameworkVersion'                    = $app.SiteConfig.NetFrameworkVersion
                            'NodeVersion'                            = $app.SiteConfig.NodeVersion
                            'NumberOfWorkers'                        = $app.SiteConfig.NumberOfWorkers
                            'PhpVersion'                             = $app.SiteConfig.PhpVersion
                            'PowerShellVersion'                      = $app.SiteConfig.PowerShellVersion
                            'PreWarmedInstanceCount'                 = $app.SiteConfig.PreWarmedInstanceCount
                            'PublicNetworkAccess'                    = $app.SiteConfig.PublicNetworkAccess
                            'PublishingUsername'                     = $app.SiteConfig.PublishingUsername
                            'Push'                                   = $app.SiteConfig.Push
                            'PythonVersion'                          = $app.SiteConfig.PythonVersion
                            'RemoteDebuggingEnabled'                 = $app.SiteConfig.RemoteDebuggingEnabled
                            'RemoteDebuggingVersion'                 = $app.SiteConfig.RemoteDebuggingVersion
                            'RequestTracingEnabled'                  = $app.SiteConfig.RequestTracingEnabled
                            'RequestTracingExpirationTime'           = $app.SiteConfig.RequestTracingExpirationTime
                            'ScmIpSecurityRestrictions'              = $app.SiteConfig.ScmIpSecurityRestrictions
                            'ScmIpSecurityRestrictionsUseMain'       = $app.SiteConfig.ScmIpSecurityRestrictionsUseMain
                            'ScmMinTlsVersion'                       = $app.SiteConfig.ScmMinTlsVersion
                            'ScmType'                                = $app.SiteConfig.ScmType
                            'TracingOptions'                         = $app.SiteConfig.TracingOptions
                            'Use32BitWorkerProcess'                  = $app.SiteConfig.Use32BitWorkerProcess
                            'VirtualApplications'                    = $app.SiteConfig.VirtualApplications
                            'VnetName'                               = $app.SiteConfig.VnetName
                            'VnetPrivatePortsCount'                  = $app.SiteConfig.VnetPrivatePortsCount
                            'VnetRouteAllEnabled'                    = $app.SiteConfig.VnetRouteAllEnabled
                            'WebsiteTimeZone'                        = $app.SiteConfig.WebsiteTimeZone
                            'WebSocketsEnabled'                      = $app.SiteConfig.WebSocketsEnabled
                            'WindowsFxVersion'                       = $app.SiteConfig.WindowsFxVersion
                            'XManagedServiceIdentityId'              = $app.SiteConfig.XManagedServiceIdentityId
                        }

                        $localWebAppSiteConfigReport.Add($webAppSiteConfigObject)

                        foreach ($appSetting in $app.SiteConfig.AppSettings) {
                            $object = [PSCustomObject]@{
                                'Subscription Id'     = $localSubscription.Id
                                'Subscription Name'   = $localSubscription.Name
                                'Resource Group Name' = $app.ResourceGroup
                                Name                  = $app.Name
                                Setting               = $appSetting.Name
                                Value                 = $appSetting.Value
                            }

                            $localWebAppSiteConfigReport.Add($object)
                        }
                    }
                }
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'App Svc Plans'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $appSvcPlanReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'WebApps'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $webAppReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'WebApps Host SSL States'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $webAppHostSslStatesReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'WebApps Site Config'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $webAppSiteConfigReport |
            Sort-Object -Property 'Subscription Name', 'Resource Group Name', Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'WebApps SiteConfig AppSettings'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $webAppSiteConfigAppSettingsReport |
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
