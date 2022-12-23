#requires -Modules ImportExcel

function New-AzReportsDefenderCloud {
    <#
    .SYNOPSIS
        Creates an Excel spreadsheet report with the details for Microsoft Defender for Cloud
    .DESCRIPTION
        Creates an Excel spreadsheet report with the details for Microsoft Defender for Cloud
    .EXAMPLE
        PS C:\> New-AzReportsDefenderCloud -Path .\temp\DefenderCloud.xlsx -Force

        Creates a report of the Microsoft Defender for Cloud and if the Path already exists it overwrites it.
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

        $scoreQuery = @'
securityresources
| where type == 'microsoft.security/securescores'
| join kind=inner (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project subscriptionId, subscriptionName = name)
    on subscriptionId
| extend percentageScore=properties.score.percentage,
    currentScore=properties.score.current,
    maxScore=properties.score.max,
    weight=properties.weight
| project subscriptionId, subscriptionName, percentageScore, currentScore, maxScore, weight
'@

        $controlsQuery = @'
securityresources
| where type == 'microsoft.security/securescores/securescorecontrols'
| join kind=inner (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project subscriptionId, subscriptionName = name)
    on subscriptionId
| extend controlName=properties.displayName,
    controlId=properties.definition.name,
    notApplicableResourceCount=properties.notApplicableResourceCount,
    unhealthyResourceCount=properties.unhealthyResourceCount,
    healthyResourceCount=properties.healthyResourceCount,
    percentageScore=properties.score.percentage,
    currentScore=properties.score.current,
    maxScore=properties.definition.properties.maxScore,
    weight=properties.weight,
    controlType=properties.definition.properties.source.sourceType
| project subscriptionId, subscriptionName, controlName, controlId, unhealthyResourceCount, healthyResourceCount, notApplicableResourceCount, percentageScore, currentScore, maxScore, weight, controlType
'@

        $recommendationsQuery = @'
securityresources
| where type == 'microsoft.security/assessments'
| join kind=inner (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project subscriptionId, subscriptionName = name)
    on subscriptionId
| extend resourceId=id,
    recommendationId=name,
    recommendationName=properties.displayName,
    source=properties.resourceDetails.Source,
    recommendationState=properties.status.code,
    description=properties.metadata.description,
    assessmentType=properties.metadata.assessmentType,
    remediationDescription=properties.metadata.remediationDescription,
    policyDefinitionId=properties.metadata.policyDefinitionId,
    implementationEffort=properties.metadata.implementationEffort,
    recommendationSeverity=properties.metadata.severity,
    category=properties.metadata.categories,
    userImpact=properties.metadata.userImpact,
    threats=properties.metadata.threats,
    portalLink=properties.links.azurePortal
| project subscriptionId, subscriptionName, resourceId, recommendationName, recommendationId, recommendationState, recommendationSeverity, description, remediationDescription, assessmentType, policyDefinitionId, implementationEffort, userImpact, category, threats, source, portalLink
'@


        $scoreResults = SearchAzGraph -Query $scoreQuery
        $controlsResults = SearchAzGraph -Query $controlsQuery
        $recommendationsResults = SearchAzGraph -Query $recommendationsQuery

        $scores = [System.Collections.ArrayList]::new()
        $controls = [System.Collections.ArrayList]::new()
        $recommendations = [System.Collections.ArrayList]::new()

        foreach ($scoreResult in $scoreResults) {
            $null = $scores.Add([PSCustomObject]@{
                    'Subscription Id'   = $scoreResult.subscriptionId
                    'Subscription Name' = $scoreResult.subscriptionName
                    'Percentage Score'  = $scoreResult.percentageScore
                    'Current Score'     = $scoreResult.currentScore
                    'Max Score'         = $scoreResult.maxScore
                    Weight              = $scoreResult.weight
                })
        }

        foreach ($controlsResult in $controlsResults) {
            $null = $controls.Add([PSCustomObject]@{
                    'Subscription Id'               = $controlsResult.subscriptionId
                    'Subscription Name'             = $controlsResult.subscriptionName
                    'Control Name'                  = $controlsResult.controlName
                    'Control Id'                    = $controlsResult.controlId
                    'Unhealthy Resource Count'      = $controlsResult.unhealthyResourceCount
                    'Healthy Resource Count'        = $controlsResult.healthyResourceCount
                    'Not Applicable Resource Count' = $controlsResult.notApplicableResourceCount
                    'Percentage Score'              = $controlsResult.percentageScore
                    'Current Score'                 = $controlsResult.currentScore
                    'Max Score'                     = $controlsResult.maxScore
                    Weight                          = $controlsResult.weight
                    'Control Type'                  = $controlsResult.controlType
                })
        }

        foreach ($recommendationsResult in $recommendationsResults) {
            $null = $recommendations.Add([PSCustomObject]@{
                    'Subscription Id'         = $recommendationsResult.subscriptionId
                    'Subscription Name'       = $recommendationsResult.subscriptionName
                    'Resource Id'             = $recommendationsResult.resourceId
                    'Recommendation Name'     = $recommendationsResult.recommendationName
                    'Recommendation Id'       = $recommendationsResult.recommendationId
                    'Recommendation State'    = $recommendationsResult.recommendationState
                    'Recommendation Severity' = $recommendationsResult.recommendationSeverity
                    Description               = $recommendationsResult.description.Replace('<br>', "`n`r").Replace('<br />', "`n`r")
                    'Remediation Description' = $recommendationsResult.remediationDescription.Replace('<br>', "`n`r").Replace('<br />', "`n`r")
                    'Assessment Type'         = $recommendationsResult.assessmentType
                    'Policy Definition Id'    = $recommendationsResult.policyDefinitionId
                    'Implementation Effort'   = $recommendationsResult.implementationEffort
                    'User Impact'             = $recommendationsResult.userImpact
                    Category                  = $recommendationsResult.category -join ', '
                    Threats                   = $recommendationsResult.threats -join ', '
                    Source                    = $recommendationsResult.source
                    'Portal Link'             = 'https://{0}' -f $recommendationsResult.portalLink
                })
        }

        $excelSplat = @{
            Path          = $Path
            WorksheetName = 'Secure Score'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $excel = $scores |
            Sort-Object -Property 'Subscription Name' |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize -HorizontalAlignment Center -NumberFormat 'Percentage'
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize -HorizontalAlignment Center

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Controls Secure Score'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $controls |
            Sort-Object -Property 'Subscription Name', 'Control Name' |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -AutoSize -HorizontalAlignment Center -NumberFormat 'Percentage'
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -HorizontalAlignment Center
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize

        $excelSplat = @{
            ExcelPackage  = $excel
            WorksheetName = 'Recommendations'
            TableStyle    = 'Medium2'
            AutoSize      = $true
            FreezeTopRow  = $true
            Style         = $excelStyle
            PassThru      = $true
        }

        $null = $recommendations |
            Sort-Object -Property 'Subscription Name', 'Resource Id', 'Recommendation Name' |
            Export-Excel @excelSplat

        $workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

        Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

        Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 3 -Width 100 -VerticalAlignment Top -WrapText
        Set-ExcelColumn -Worksheet $workSheet -Column 4 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 5 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 6 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 7 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 8 -Width 100 -WrapText -VerticalAlignment Top -NumberFormat 'Text'
        Set-ExcelColumn -Worksheet $workSheet -Column 9 -Width 100 -WrapText -VerticalAlignment Top -NumberFormat 'Text'
        Set-ExcelColumn -Worksheet $workSheet -Column 10 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 11 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 12 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 13 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 14 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 15 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 16 -AutoSize -VerticalAlignment Top
        Set-ExcelColumn -Worksheet $workSheet -Column 17 -AutoSize -VerticalAlignment Top

        if ($NoInvoke) {
            Close-ExcelPackage -ExcelPackage $excel
        } else {
            Close-ExcelPackage -ExcelPackage $excel -Show
        }
    } catch {
        throw $PSItem
    }
}
