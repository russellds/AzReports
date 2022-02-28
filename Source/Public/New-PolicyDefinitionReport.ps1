#requires -Modules ImportExcel

function New-PolicyDefinitionReport {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory)]
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'BuiltIn')]
        [Parameter(ParameterSetName = 'Custom')]
        [System.IO.FileInfo]
        $Path,

        [Parameter(ParameterSetName = 'BuiltIn')]
        [switch]
        $BuiltIn,

        [Parameter(ParameterSetName = 'Custom')]
        [switch]
        $Custom,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'BuiltIn')]
        [Parameter(ParameterSetName = 'Custom')]
        [switch]
        $NoInvoke,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'BuiltIn')]
        [Parameter(ParameterSetName = 'Custom')]
        [switch]
        $Force
    )

    $PSBoundParameters.Keys

    if ($Path.Extension -ne '.xlsx') {
        throw 'File extension must be .xlsx!'
    }

    if (Test-Path -Path $Path.DirectoryName) {
        if (Test-Path -Path $Path.FullName) {
            if ($Force) {
                [void](Remove-Item -Path $Path.FullName -Force)
            } else {
                throw "$( $Path.FullName ) already exists, pass -Force to overwrite!"
            }
        }
    } else {
        [void](New-Item -Path $Path.DirectoryName -ItemType Directory -Force)
    }

    if ($PSBoundParameters.Keys -contains 'BuiltIn') {
        $policyDefinitions = Get-AzPolicyDefinition -Builtin
    } elseif ($PSBoundParameters.Keys -contains 'Custom') {
        $policyDefinitions = Get-AzPolicyDefinition -Custom
    } else {
        $policyDefinitions = Get-AzPolicyDefinition
    }

    $objects = @()

    foreach ($policyDefinition in $policyDefinitions) {
        Write-Verbose "Policy Name: $($policyDefinition.Properties.DisplayName)" -Verbose

        $item = [ordered]@{
            Name = $policyDefinition.Name
            Category = $policyDefinition.Properties.Metadata.category
            Type = $policyDefinition.Properties.PolicyType
            'Display Name' = $policyDefinition.Properties.DisplayName
            Description = $policyDefinition.Properties.Description
            'Resource Id' = $policyDefinition.ResourceId
            'Available Effects' = $policyDefinition.Properties.PolicyRule.then.effect
            'Parameter Name' = $null
            'Parameter Type' = $null
            'Parameter Display Name' = $null
            'Parameter Description' = $null
            'Allowed Values' = $null
            'Default Value' = $null
            'Desired Value' = $null
        }

        if ($policyDefinition.Properties.Parameters) {
            $parameters = $policyDefinition.Properties.Parameters.PSObject.Members |
                Where-Object { $_.MemberType -eq 'NoteProperty' } |
                Select-Object -ExpandProperty Name

            foreach ($parameter in $parameters) {
                Write-Verbose "Parameter Name: $parameter" -Verbose
                $item.'Parameter Name' = $parameter
                $item.'Parameter Type' = $policyDefinition.Properties.Parameters.$parameter.type
                $item.'Parameter Display Name' = $policyDefinition.Properties.Parameters.$parameter.metadata.displayName
                $item.'Parameter Description' = $policyDefinition.Properties.Parameters.$parameter.metadata.description
                $item.'Allowed Values' = $policyDefinition.Properties.Parameters.$parameter.allowedValues -join ', '

                if ($policyDefinition.Properties.Parameters.$parameter.type -eq 'Object') {
                    $item.'Default Value' = $policyDefinition.Properties.Parameters.$parameter.defaultValue |
                        ConvertTo-Json -Compress
                } else {
                    $item.'Default Value' = $policyDefinition.Properties.Parameters.$parameter.defaultValue -join ', '
                }

                $objects += [PSCustomObject]$item
            }

        } else {
            $objects += [PSCustomObject]$item
        }
    }

    $excelStyle = New-ExcelStyle -VerticalAlignment Top

    $excelSplat = @{
        Path = $Path
        WorksheetName = 'Policies'
        TableStyle = 'Medium2'
        AutoSize = $true
        FreezeTopRow = $true
        Style = $excelStyle
        PassThru = $true
    }

    $excel = $objects |
        Sort-Object -Property Category, 'Display Name', 'Parameter Name' |
        Export-Excel @excelSplat

    $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

    Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

    Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize -Hide
    Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 4 -Width 50 -WrapText
    Set-ExcelColumn -Worksheet $workSheet -Column 5 -Width 60 -WrapText
    Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize -HorizontalAlignment Left
    Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 8 -Width 50 -WrapText -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 10 -Width 50 -WrapText
    Set-ExcelColumn -Worksheet $workSheet -Column 11 -Width 50 -WrapText
    Set-ExcelColumn -Worksheet $workSheet -Column 12 -Width 50 -WrapText -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 13 -Width 50 -WrapText -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 14 -Width 50 -WrapText -HorizontalAlignment Center

    if ($NoInvoke) {
        Close-ExcelPackage -ExcelPackage $excel
    } else {
        Close-ExcelPackage -ExcelPackage $excel -Show
    }
}
