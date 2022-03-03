#requires -Modules ImportExcel

function New-PolicyAssignmentReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        [Parameter(Mandatory)]
        [string]
        $Name,

        [switch]
        $NoInvoke,

        [switch]
        $Force
    )
    $InformationPreference = 'Continue'

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

    $policyAssignment = Get-AzPolicyAssignment -Name SecurityCenterBuiltIn

    $policySetDefinition = Get-AzPolicySetDefinition -ResourceId $policyAssignment.Properties.PolicyDefinitionId

    $objects = @()

    foreach ($policySetPolicyDefinition in $policySetDefinition.Properties.PolicyDefinitions) {
        $parameters = $policySetPolicyDefinition.parameters.PSObject.Members |
            Where-Object { $_.MemberType -eq 'NoteProperty' } |
            Select-Object -ExpandProperty Name

        $policyDefinition = Get-AzPolicyDefinition -ResourceId $policySetPolicyDefinition.policyDefinitionId

        foreach ($parameter in $parameters) {
            Write-Information "Policy: $($policyDefinition.Properties.DisplayName) Parameter: $parameter"
            if ($policySetPolicyDefinition.parameters.$parameter.value.GetType().Name -eq 'String') {
                if ($policySetPolicyDefinition.parameters.$parameter.value.StartsWith('[parameters')) {
                    $item = [PSCustomObject]@{
                        Name = $policyDefinition.Properties.DisplayName
                        Category = $policyDefinition.Properties.Metadata.category
                        'Parameter Name' = $($policySetPolicyDefinition.parameters.$parameter.value.Split("'"))[1]
                        'Parameter Type' = $policyDefinition.Properties.Parameters.$parameter.type
                        'Allowed Values' = $policyDefinition.Properties.Parameters.$parameter.allowedValues -join ', '
                        'Default Value' = $null
                        'Desired Value' = $null
                    }

                    if ($policyDefinition.Properties.Parameters.$parameter.type -eq 'Object') {
                        $item.'Default Value' = $policyDefinition.Properties.Parameters.$parameter.defaultValue |
                            ConvertTo-Json -Compress
                    } else {
                        $item.'Default Value' = $policyDefinition.Properties.Parameters.$parameter.defaultValue -join ', '
                    }

                    $objects += $item
                }
            }
        }
    }

    $excelSplat = @{
        Path = $Path
        WorksheetName = $Name
        TableStyle = 'Medium2'
        AutoSize = $true
        FreezeTopRow = $true
        Style = $excelStyle
        PassThru = $true
    }

    $excel = $objects |
        Sort-Object -Property Category, Name |
        Export-Excel @excelSplat

    $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

    Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

    Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
    Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
    Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
    Set-ExcelColumn -Worksheet $workSheet -Column 4 -Width 19 -HorizontalAlignment Center
    Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize
    Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize -HorizontalAlignment Right
    Set-ExcelColumn -Worksheet $workSheet -Column 7 -Width 18


    if ($NoInvoke) {
        Close-ExcelPackage -ExcelPackage $excel
    } else {
        Close-ExcelPackage -ExcelPackage $excel -Show
    }

}
