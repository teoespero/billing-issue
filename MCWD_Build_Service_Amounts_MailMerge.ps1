# Author: Teo Espero
# Revision 19 - Bulk Excel transfers and visible progress, October 5, 2026.
# Requires Windows and desktop Excel. SAMPLE amounts are for testing only.
# Example: .\MCWD_Build_Service_Amounts_MailMerge.ps1
[CmdletBinding()]
param([string]$WorkbookPath = (Join-Path $PSScriptRoot 'MCWD_Service_Amounts_SAMPLE.xlsx'))
$ErrorActionPreference = 'Stop'
$excel = $null; $source = $null
function Headers($values) {
    $map = @{}
    for ($c=1; $c -le $values.GetLength(1); $c++) {
        $name = [string]$values[1,$c]
        if ($name) { $map[$name] = $c }
    }
    return $map
}
function Cell($values,$row,$map,$name) {
    if (-not $map.ContainsKey($name)) { throw "Missing column: $name" }
    return $values[$row,$map[$name]]
}
function Money($value) {
    return ([decimal]$value).ToString('$#,##0.00;-$#,##0.00',[cultureinfo]'en-US')
}
try {
    $WorkbookPath = (Resolve-Path $WorkbookPath).Path
    Write-Host 'Opening workbook in Excel...' -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $source = $excel.Workbooks.Open($WorkbookPath,0,$false)
    if ($source.ReadOnly) { throw 'Close the workbook in Excel or Word before refreshing.' }
    Write-Host 'Reading ServiceAmounts in one batch...' -ForegroundColor Cyan
    $entry = $source.Worksheets.Item('ServiceAmounts')
    $used = $entry.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1
    $inputRange = $entry.Range($entry.Cells.Item(1,1),$entry.Cells.Item($lastRow,$lastCol))
    $inputValues = $inputRange.Value2
    if ($lastRow -lt 2) { throw 'ServiceAmounts contains no data rows.' }
    $eh = Headers $inputValues
    foreach ($name in @('AccountNumber','CustomerName','AddressLine1','AddressLine2','CityStateZIP','Service','BillingPeriod','OriginalServiceAmount','CorrectedServiceAmount')) {
        if (-not $eh.ContainsKey($name)) { throw "Missing column: $name" }
    }
    Write-Host "Validating $($lastRow-1) rows and grouping accounts..." -ForegroundColor Cyan
    $accounts = [ordered]@{}
    $seen = @{}
    for ($r=2; $r -le $lastRow; $r++) {
        if ($r % 100 -eq 0 -or $r -eq $lastRow) {
            Write-Progress -Id 1 -Activity 'Preparing mail merge' -Status "Validating row $r of $lastRow" -PercentComplete ([int](70*$r/$lastRow))
        }
        $account=[string](Cell $inputValues $r $eh 'AccountNumber')
        if (-not $account) { continue }
        if (-not $accounts.Contains($account)) {
            $accounts[$account] = @{
                Name=[string](Cell $inputValues $r $eh 'CustomerName')
                Address1=[string](Cell $inputValues $r $eh 'AddressLine1')
                Address2=[string](Cell $inputValues $r $eh 'AddressLine2')
                City=[string](Cell $inputValues $r $eh 'CityStateZIP')
                Lines=New-Object 'System.Collections.Generic.List[string]'
                Periods=New-Object 'System.Collections.Generic.List[string]'
                Total=[decimal]0
            }
        }
        $service=[string](Cell $inputValues $r $eh 'Service')
        $period=[string](Cell $inputValues $r $eh 'BillingPeriod')
        if (-not $service -or -not $period) { throw "Missing service or period at row $r." }
        $key="$account|$service|$period"
        if ($seen.ContainsKey($key)) { throw "Duplicate account/service/month: $key" }
        $seen[$key]=$true
        $original=Cell $inputValues $r $eh 'OriginalServiceAmount'
        $corrected=Cell $inputValues $r $eh 'CorrectedServiceAmount'
        if ([string]::IsNullOrWhiteSpace([string]$original) -or [string]::IsNullOrWhiteSpace([string]$corrected)) {
            throw "Enter both service amounts at row $r before merging."
        }
        $delta=[math]::Round(([decimal]$corrected-[decimal]$original),2)
        if ($eh.ContainsKey('KnownAdjustment')) {
            $known=Cell $inputValues $r $eh 'KnownAdjustment'
            if (-not [string]::IsNullOrWhiteSpace([string]$known) -and [math]::Abs($delta-[decimal]$known) -gt 0.005) {
                throw "Amounts do not match KnownAdjustment at row $r."
            }
        }
        if ($delta -eq 0) { continue }
        $rec=$accounts[$account]
        # Literal tabs align all five columns in the Word summary.
        $rec.Lines.Add((@($service,$period,(Money $original),(Money $corrected),(Money $delta)) -join "`t"))
        if (-not $rec.Periods.Contains($period)) { $rec.Periods.Add($period) }
        $rec.Total += $delta
    }
    # Assemble the output in memory before changing the merge sheet.
    $names=@('CustomerName','AccountNumber','AddressLine1','AddressLine2','CityStateZIP','AffectedBillingPeriods','SummaryDetailLines','TotalAdjustment')
    $recordCount=0
    foreach ($account in $accounts.Keys) { if ($accounts[$account].Lines.Count -gt 0) { $recordCount++ } }
    $outputValues=[Array]::CreateInstance([object], [int[]]@(($recordCount+1),8))
    for ($c=0;$c -lt 8;$c++) { $outputValues[0,$c]=$names[$c] }
    $row=1
    foreach ($account in $accounts.Keys) {
        $rec=$accounts[$account]
        if ($rec.Lines.Count -eq 0) { continue }
        $periodText=$rec.Periods -join ' and '
        if ($rec.Periods.Count -eq 2 -and $rec.Periods.Contains('July 2026') -and $rec.Periods.Contains('August 2026')) { $periodText='July and August 2026' }
        $values=@($rec.Name,$account,$rec.Address1,$rec.Address2,$rec.City,$periodText,($rec.Lines -join "`n"))
        for ($c=0;$c -lt 7;$c++) { $outputValues[$row,$c]=[string]$values[$c] }
        $outputValues[$row,7]=[double]$rec.Total
        $row++
    }
    Write-Host "Writing $recordCount account records in one batch..." -ForegroundColor Cyan
    Write-Progress -Id 1 -Activity 'Preparing mail merge' -Status 'Writing MailMergeData' -PercentComplete 80
    $merge=$null
    foreach ($sheet in $source.Worksheets) { if ($sheet.Name -eq 'MailMergeData') { $merge=$sheet; break } }
    if ($null -eq $merge) { $merge=$source.Worksheets.Add(); $merge.Name='MailMergeData' }
    $merge.UsedRange.Clear() | Out-Null
    $merge.Columns.Item(2).NumberFormat='@'
    $merge.Columns.Item(7).NumberFormat='@'
    $outputRange=$merge.Range("A1:H$($recordCount+1)")
    $outputRange.Value2=$outputValues
    $merge.Columns.Item(8).NumberFormat='$#,##0.00;-$#,##0.00'
    $merge.Rows.Item(1).Font.Bold=$true
    $merge.Range('A:F').ColumnWidth=28
    $merge.Columns.Item(8).ColumnWidth=20
    $merge.Columns.Item(7).ColumnWidth=90; $merge.Columns.Item(7).WrapText=$true
    Write-Host 'Saving workbook...' -ForegroundColor Cyan
    Write-Progress -Id 1 -Activity 'Preparing mail merge' -Status 'Saving workbook' -PercentComplete 95
    $source.Save()
    Write-Host "Created $recordCount account records: $WorkbookPath"
    Write-Host 'In Word, select MailMergeData as the recipient sheet. SAMPLE amounts are fictional.'
}
finally {
    Write-Progress -Id 1 -Activity 'Preparing mail merge' -Completed
    if ($null -ne $source) { $source.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
