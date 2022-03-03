function New-AzReportsRegion {
    <#
    .SYNOPSIS
        Creates an Markdown file report with the details for Azure Policy Assignment
    .DESCRIPTION
        Creates an Markdown file report with the details for Azure Policy Assignment
    .EXAMPLE
        PS C:\> New-AzReportsPolicyAssignment -Path .\temp\SecurityCenterBuiltIn.xlsx -Name SecurityCenterBuiltIn -Force

        Creates a report of the Azure Policy Assignment and if the Path already exists it overwrites it.
    .INPUTS
        None
    .OUTPUTS
        Markdown File
    #>
    [CmdletBinding()]
    param(
        # Path to create the Markdown report. Must end with '.md'.
        [System.IO.FileInfo]
        $Path,

        # Do not automatically open the generated Markdown file.
        [switch]
        $NoInvoke,

        # Overwrite existing Markdown file.
        [switch]
        $Force
    )
    CheckAzContext

    if ($Path) {
        CheckPath -Path $Path -Extension '.md' -Force:$Force
    }

    $azRegions = Get-AzLocation |
        Sort-Object -Property DisplayName

    $customAzRegions = [System.Collections.ArrayList]@()

    foreach ($azRegion in $azRegions) {
        $shortLocation = ''
        $shortLocationSubstitions = @(
            @{
                Southeast = 'SE'
            }
        )
        $displayNameComponents = $azRegion.DisplayName.Split(' ')

        foreach ($displayNameComponent in $displayNameComponents) {
            if ($shortLocationSubstitions.Keys -contains $displayNameComponent) {
                $shortLocation += $shortLocationSubstitions.$displayNameComponent
            } elseif ([Int]::TryParse($displayNameComponent, [ref]$null)) {
                $shortLocation += $displayNameComponent
            } else {
                $shortLocation += $displayNameComponent -creplace '[^A-Z]'
            }
        }

        $customAzRegions += [PSCustomObject]@{
            'Display Name'        = $azRegion.DisplayName
            Region                = $azRegion.Location
            'Region Length'       = $azRegion.Location.Length
            'Short Region'        = $shortLocation.ToLower()
            'Short Region Length' = $shortLocation.Length
        }
    }

    $headers = $customAzRegions |
        Get-Member -MemberType NoteProperty |
        Select-Object -ExpandProperty Name

    $sbTable = [System.Text.StringBuilder]'|'

    $headerPadding = @{}

    foreach ($header in $headers) {
        $headerValueLengths = $customAzRegions.'Display Name' |
            Select-Object -ExpandProperty Length -Unique |
            Sort-Object -Descending

        if ($header.Length -gt $headerValueLengths[0]) {
            $headerPadding.$header = $header.Length
        } else {
            $headerPadding.$header = $headerValueLengths[0]
        }

        $null = $sbTable.Append((' {0} |' -f $header.PadRight($headerPadding.$header, ' ')))
    }

    $null = $sbTable.AppendLine('')
    $null = $sbTable.Append('|')

    foreach ($header in $headers) {
        $null = $sbTable.Append((' {0} |' -f ''.PadRight($headerPadding.$header, '-')))
    }

    $null = $sbTable.AppendLine('')

    foreach ($customAzRegion in $customAzRegions) {
        $null = $sbTable.Append('|')

        foreach ($header in $headers) {
            $null = $sbTable.Append((' {0} |' -f $customAzRegion.$header.ToString().PadRight($headerPadding.$header, ' ')))
        }

        $null = $sbTable.AppendLine('')
    }

    if ($Path) {
        $sbTable.ToString() |
            Out-File -FilePath $Path.FullName

        if (-not $NoInvoke) {
            Invoke-Item -Path $Path.FullName
        }
    } else {
        $sbTable.ToString()
    }
}
