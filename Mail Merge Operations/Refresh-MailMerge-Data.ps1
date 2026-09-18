param(
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function To-Text {
    param($Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function To-Number {
    param($Value)
    if ($null -eq $Value -or $Value -eq "") { return 0.0 }
    try { return [double]$Value } catch { return 0.0 }
}

function Service-Label {
    param([string]$Service)
    $s = $Service.Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    if ($s -match '(?i)\bincrease$') { return $s }
    return "$s increase"
}

function Write-ExcelColumn {
    param(
        $Worksheet,
        [int]$StartRow,
        [int]$ColumnNumber,
        [object[]]$Values
    )

    $count = $Values.Count
    if ($count -eq 0) { return }

    # Excel COM is more reliable when each column is written as its own
    # two-dimensional N x 1 SAFEARRAY.
    $arr = [System.Array]::CreateInstance([object], @($count, 1), @(0, 0))

    for ($i = 0; $i -lt $count; $i++) {
        $arr.SetValue($Values[$i], $i, 0)
    }

    $endRow = $StartRow + $count - 1
    $target = $Worksheet.Range(
        $Worksheet.Cells.Item($StartRow, $ColumnNumber),
        $Worksheet.Cells.Item($endRow, $ColumnNumber)
    )

    $target.Value2 = $arr
    Release-ComObject $target
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $preferred = Join-Path $scriptDir "2026-0918-Customer-Billing-Adjustment-Row-Based-Static-Mail-Merge-Data-TE.xlsx"

    if (Test-Path $preferred) {
        $WorkbookPath = $preferred
    }
    else {
        $candidates = @(Get-ChildItem -Path $scriptDir -Filter "*.xlsx" -File |
            Where-Object { $_.Name -notlike "~$*" })

        if ($candidates.Count -eq 1) {
            $WorkbookPath = $candidates[0].FullName
        }
        elseif ($candidates.Count -eq 0) {
            throw "No .xlsx workbook was found in $scriptDir."
        }
        else {
            throw "More than one .xlsx workbook was found. Run:`n  .\Refresh-MailMerge-Data-v5-ColumnWrite.ps1 -WorkbookPath `"C:\path\file.xlsx`""
        }
    }
}

$WorkbookPath = (Resolve-Path $WorkbookPath).Path

$excel = $null
$workbook = $null
$sourceSheet = $null
$mergeSheet = $null
$sourceRange = $null
$existingUsed = $null

try {
    Write-Host "Opening workbook..." -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    # Do not change Excel.Calculation.
    $workbook = $excel.Workbooks.Open($WorkbookPath)

    $sourceSheet = $workbook.Worksheets.Item("AdjustmentData")
    $mergeSheet = $workbook.Worksheets.Item("MailMergeData")

    # Real last row based on AccountNumber in column B.
    $xlUp = -4162
    $lastRow = $sourceSheet.Cells.Item($sourceSheet.Rows.Count, 2).End($xlUp).Row

    if ($lastRow -lt 2) {
        throw "AdjustmentData does not contain any account rows."
    }

    $dataRowCount = $lastRow - 1
    Write-Host "Actual data rows found: $dataRowCount" -ForegroundColor Cyan
    Write-Host "Loading AdjustmentData into memory..." -ForegroundColor Cyan

    $sourceRange = $sourceSheet.Range("A2:I$lastRow")
    $data = $sourceRange.Value2

    $accounts = [ordered]@{}

    # Handle either a 2-D array (multiple rows) or a scalar row.
    if ($data -is [System.Array] -and $data.Rank -eq 2) {
        $rowLB = $data.GetLowerBound(0)
        $rowUB = $data.GetUpperBound(0)
        $colLB = $data.GetLowerBound(1)

        for ($i = $rowLB; $i -le $rowUB; $i++) {
            $customer     = To-Text   $data.GetValue($i, $colLB + 0)
            $account      = To-Text   $data.GetValue($i, $colLB + 1)
            $address      = To-Text   $data.GetValue($i, $colLB + 2)
            $cityStateZip = To-Text   $data.GetValue($i, $colLB + 3)
            $period       = To-Text   $data.GetValue($i, $colLB + 4)
            $service      = To-Text   $data.GetValue($i, $colLB + 5)
            $difference   = To-Number $data.GetValue($i, $colLB + 6)
            $original     = To-Number $data.GetValue($i, $colLB + 7)
            $corrected    = To-Number $data.GetValue($i, $colLB + 8)

            if ([string]::IsNullOrWhiteSpace($account)) { continue }

            if (-not $accounts.Contains($account)) {
                $accounts[$account] = [ordered]@{
                    CustomerName        = $customer
                    AccountNumber       = $account
                    AddressLine1        = $address
                    CityStateZIP        = $cityStateZip

                    HasJuly             = $false
                    JulyOriginalBill    = 0.0
                    JulyCorrectedBill   = 0.0
                    JulyServices        = New-Object System.Collections.Generic.List[string]
                    JulyDifferences     = New-Object System.Collections.Generic.List[double]
                    JulyTotalIncrease   = 0.0

                    HasAugust           = $false
                    AugustOriginalBill  = 0.0
                    AugustCorrectedBill = 0.0
                    AugustServices      = New-Object System.Collections.Generic.List[string]
                    AugustDifferences   = New-Object System.Collections.Generic.List[double]
                    AugustTotalIncrease = 0.0

                    TotalAdjustment     = 0.0
                }
            }

            $rec = $accounts[$account]

            if ([string]::IsNullOrWhiteSpace($rec.CustomerName) -and $customer) {
                $rec.CustomerName = $customer
            }
            if ([string]::IsNullOrWhiteSpace($rec.AddressLine1) -and $address) {
                $rec.AddressLine1 = $address
            }
            if ([string]::IsNullOrWhiteSpace($rec.CityStateZIP) -and $cityStateZip) {
                $rec.CityStateZIP = $cityStateZip
            }

            if ($period -eq "July 2026") {
                $rec.HasJuly = $true
                if ($original -ne 0)  { $rec.JulyOriginalBill = $original }
                if ($corrected -ne 0) { $rec.JulyCorrectedBill = $corrected }

                if (-not [string]::IsNullOrWhiteSpace($service)) {
                    $rec.JulyServices.Add((Service-Label $service))
                    $rec.JulyDifferences.Add($difference)
                }
                $rec.JulyTotalIncrease += $difference
            }
            elseif ($period -eq "August 2026") {
                $rec.HasAugust = $true
                if ($original -ne 0)  { $rec.AugustOriginalBill = $original }
                if ($corrected -ne 0) { $rec.AugustCorrectedBill = $corrected }

                if (-not [string]::IsNullOrWhiteSpace($service)) {
                    $rec.AugustServices.Add((Service-Label $service))
                    $rec.AugustDifferences.Add($difference)
                }
                $rec.AugustTotalIncrease += $difference
            }

            $rec.TotalAdjustment += $difference
        }
    }
    else {
        throw "The source range could not be read as a table."
    }

    Write-Host ("Customer records created in memory: {0}" -f $accounts.Count) -ForegroundColor Cyan
    Write-Host "Clearing old MailMergeData..." -ForegroundColor Cyan

    $existingUsed = $mergeSheet.UsedRange
    $existingLastRow = $existingUsed.Row + $existingUsed.Rows.Count - 1
    if ($existingLastRow -ge 2) {
        $mergeSheet.Range("A2:P$existingLastRow").ClearContents() | Out-Null
    }

    if ($accounts.Count -eq 0) {
        throw "No usable account records were found."
    }

    Write-Host "Building output columns..." -ForegroundColor Cyan

    # Create 16 independent output columns. Writing one Excel column at a time
    # avoids the COM multidimensional-array issue that populated only column A.
    $cols = @()
    for ($c = 0; $c -lt 16; $c++) {
        $cols += ,(New-Object System.Collections.Generic.List[object])
    }

    foreach ($account in $accounts.Keys) {
        $rec = $accounts[$account]

        if ($rec.HasJuly -and $rec.HasAugust) {
            $affectedPeriods = "July and August 2026"
        }
        elseif ($rec.HasJuly) {
            $affectedPeriods = "July 2026"
        }
        elseif ($rec.HasAugust) {
            $affectedPeriods = "August 2026"
        }
        else {
            $affectedPeriods = ""
        }

        $julyServiceLines = [string]::Join("`n", $rec.JulyServices)
        $julyDifferenceLines = [string]::Join("`n", @(
            $rec.JulyDifferences | ForEach-Object { '$' + ('{0:N2}' -f $_) }
        ))

        $augustServiceLines = [string]::Join("`n", $rec.AugustServices)
        $augustDifferenceLines = [string]::Join("`n", @(
            $rec.AugustDifferences | ForEach-Object { '$' + ('{0:N2}' -f $_) }
        ))

        $cols[0].Add([string]$rec.CustomerName)
        $cols[1].Add([string]$rec.AccountNumber)
        $cols[2].Add([string]$rec.AddressLine1)
        $cols[3].Add([string]$rec.CityStateZIP)
        $cols[4].Add([string]$affectedPeriods)
        $cols[5].Add($(if ($rec.HasJuly)   { [double]$rec.JulyOriginalBill } else { $null }))
        $cols[6].Add($(if ($rec.HasJuly)   { [double]$rec.JulyCorrectedBill } else { $null }))
        $cols[7].Add($(if ($rec.HasAugust) { [double]$rec.AugustOriginalBill } else { $null }))
        $cols[8].Add($(if ($rec.HasAugust) { [double]$rec.AugustCorrectedBill } else { $null }))
        $cols[9].Add([string]$julyServiceLines)
        $cols[10].Add([string]$julyDifferenceLines)
        $cols[11].Add([double]$rec.JulyTotalIncrease)
        $cols[12].Add([string]$augustServiceLines)
        $cols[13].Add([string]$augustDifferenceLines)
        $cols[14].Add([double]$rec.AugustTotalIncrease)
        $cols[15].Add([double]$rec.TotalAdjustment)
    }

    Write-Host "Writing all 16 MailMergeData columns..." -ForegroundColor Cyan

    for ($c = 0; $c -lt 16; $c++) {
        Write-ExcelColumn -Worksheet $mergeSheet -StartRow 2 -ColumnNumber ($c + 1) -Values $cols[$c].ToArray()
        Write-Host ("  Wrote column {0} of 16" -f ($c + 1)) -ForegroundColor DarkCyan
    }

    $lastOutRow = $accounts.Count + 1

    Write-Host "Formatting MailMergeData..." -ForegroundColor Cyan

    $mergeSheet.Range("B2:B$lastOutRow").NumberFormat = "@"
    $mergeSheet.Range("F2:I$lastOutRow").NumberFormat = '$#,##0.00'
    $mergeSheet.Range("L2:L$lastOutRow").NumberFormat = '$#,##0.00'
    $mergeSheet.Range("O2:P$lastOutRow").NumberFormat = '$#,##0.00'
    $mergeSheet.Range("J2:N$lastOutRow").WrapText = $true
    $mergeSheet.Range("A2:P$lastOutRow").VerticalAlignment = -4160

    Write-Host "Saving workbook..." -ForegroundColor Cyan
    $workbook.Save()

    Write-Host ""
    Write-Host "MailMergeData refreshed successfully." -ForegroundColor Green
    Write-Host ("Source service rows: {0}" -f $dataRowCount) -ForegroundColor Green
    Write-Host ("Customer records:   {0}" -f $accounts.Count) -ForegroundColor Green
    Write-Host ("Workbook: {0}" -f $WorkbookPath) -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Refresh failed." -ForegroundColor Red
    Write-Host ("Message: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if ($_.InvocationInfo) {
        Write-Host ("Script line: {0}" -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Red
        Write-Host ("Command: {0}" -f $_.InvocationInfo.Line) -ForegroundColor Red
    }
    throw
}
finally {
    if ($sourceRange) { Release-ComObject $sourceRange }
    if ($existingUsed) { Release-ComObject $existingUsed }
    if ($sourceSheet) { Release-ComObject $sourceSheet }
    if ($mergeSheet) { Release-ComObject $mergeSheet }

    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }

    if ($excel) {
        try { $excel.EnableEvents = $true } catch {}
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
