#requires -Modules ImportExcel

function New-AzReportsApplicationGateway {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Azure Application Gateway
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Azure Application Gateway
    .EXAMPLE
        PS C:\> New-AzReportsApplicationGateway -Path .\temp\ApplicationGateway.xlsx -Force

        Creates a report of the Azure Application Gateway and if the Path already exists it overwrites it.
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

        $appGateways = Get-AzApplicationGateway

        $customAppGatewayProperties = @(
            'ResourceGroupName',
            'Name',
            # 'AuthenticationCertificates', Maybe seperate worksheet.
            # 'AuthenticationCertificatesText', Probably not include.
            'AutoscaleConfiguration',
            # 'BackendAddressPools', Seperate worksheet.
            # 'BackendAddressPoolsText', Seperate worksheet.
            # 'BackendHttpSettingsCollection', Seperate worksheet.
            # 'BackendHttpSettingsCollectionText', Seperate worksheet.
            # 'CustomErrorConfigurations', Seperate worksheet.
            'EnableFips',
            'EnableHttp2',
            'Etag',
            'FirewallPolicy',
            'FirewallPolicyText',
            'ForceFirewallPolicyAssociation',
            # 'FrontendIPConfigurations', Seperate Worksheet
            # 'FrontendIpConfigurationsText', Seperate Worksheet
            # 'FrontendPorts', Seperate Worksheet
            # 'FrontendPortsText', Seperate Worksheet
            # 'GatewayIPConfigurations', Seperate Worksheet
            # 'GatewayIpConfigurationsText', Seperate Worksheet
            # 'HttpListeners', Seperate Worksheet
            # 'HttpListenersText', Seperate Worksheet
            'Identity',
            'IdentityText',
            'Location',
            'OperationalState',
            # 'PrivateEndpointConnections', Seperate worksheet.
            # 'PrivateLinkConfigurations', Seperate worksheet.
            # 'PrivateLinkConfigurationsText', Seperate worksheet.
            # 'PrivateLinkEndpointConnectionsText', Seperate worksheet.
            # 'Probes', Seperate worksheet.
            # 'ProbesText', Seperate worksheet.
            'ProvisioningState',
            # 'RedirectConfigurations', Seperate worksheet.
            # 'RequestRoutingRules', Seperate worksheet.
            # 'RequestRoutingRulesText', Seperate worksheet.
            'ResourceGuid',
            # 'RewriteRuleSets', Seperate worksheet.
            # 'RewriteRuleSetsText', Seperate worksheet.
            'Sku', # Object
            # 'SslCertificates', Seperate worksheet.
            # 'SslCertificatesText', Seperate worksheet.
            # 'SslPolicy', Seperate worksheet.
            # 'SslPolicyText', Seperate worksheet.
            # 'SslProfiles', Seperate worksheet.
            # 'SslProfilesText', Seperate worksheet.
            'Tag',
            'TagsTable',
            # 'TrustedClientCertificates', Seperate worksheet.
            # 'TrustedClientCertificatesText', Seperate worksheet.
            # 'TrustedRootCertificates', Seperate worksheet.
            'Type',
            # 'UrlPathMaps', Seperate worksheet.
            # 'UrlPathMapsText', Seperate worksheet.
            'WebApplicationFirewallConfiguration',
            'Zones'
        )

        $customAppGatewayObjects = $appGateways |
            Select-Object -Property $customAppGatewayProperties

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'Profiles'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $customAppGatewayObjects |
            Sort-Object -Property ResourecGroupName, Name |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        # Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Endpoints'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        # $customEndpointObjects |
        #     Sort-Object -Property ResourecGroupName, ProfileName, Name |
        #     Export-Excel @excelSplat

        # $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        # Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        # Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -HorizontalAlignment Center
        # Set-ExcelColumn -Worksheet $workSheet -Column 14 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 15 -AutoSize
        # Set-ExcelColumn -Worksheet $workSheet -Column 16 -AutoSize

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
