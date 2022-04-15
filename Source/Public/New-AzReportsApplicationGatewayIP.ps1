$resourceGroups = Get-AzResourceGroup

$objects = @()

foreach ($resourceGroup in $resourceGroups) {
    $webApps = Get-AzWebApp -ResourceGroupName $resourceGroup.ResourceGroupName

    if ($webApps) {
        foreach ($webApp in $webApps) {
            $objects += [PSCustomObject]@{
                ResourceGroup = $webApp.ResourceGroup
                Name = $webApp.Name
                Location = $webApp.Location
                OutboundIpAddresses = ($webApp.OutboundIpAddresses.Split(',')) -join ', '
                PossibleOutboundIpAddresses = ($webApp.PossibleOutboundIpAddresses.Split(',')) -join ', '
            }
        }
    }
}

$path = [System.IO.FileInfo]'C:\temp\Sync1WebAppIpAddressReport.xlsx'

if (Test-Path -Path $path.FullName) {
    $null = Remove-Item -Path $path.FullName -Force
}

$excelStyle = New-ExcelStyle -VerticalAlignment Top

$excelSplat = @{
    Path = $path.FullName
    WorksheetName = 'IpAddresses'
    TableStyle = 'Medium2'
    AutoSize = $true
    FreezeTopRow = $true
    Style = $excelStyle
    PassThru = $true
}

$excel = $objects |
    Sort-Object -Property ResourceGroup, Name |
    Export-Excel @excelSplat

$workSheet = $excel.Workbook.Worksheets[$excelSplat.WorksheetName]

Set-ExcelRow -Worksheet $workSheet -Row 1 -Bold -HorizontalAlignment Center

Set-ExcelColumn -Worksheet $workSheet -Column 1 -AutoSize
Set-ExcelColumn -Worksheet $workSheet -Column 2 -AutoSize
Set-ExcelColumn -Worksheet $workSheet -Column 3 -AutoSize
Set-ExcelColumn -Worksheet $workSheet -Column 4 -Width 50 -WrapText
Set-ExcelColumn -Worksheet $workSheet -Column 5 -Width 50 -WrapText

Close-ExcelPackage -ExcelPackage $excel -Show
