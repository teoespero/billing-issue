# Customer Billing Adjustment Mail Merge Refresh Script

## File
`Refresh-MailMerge-Data-v6-Exclude-Zero.ps1`

## Purpose
This PowerShell script prepares customer billing adjustment data for a Microsoft Word mail merge.

The source Excel data is stored in a row-by-row format, where each affected service for an account is entered as a separate row. The script groups those rows by customer account and creates a separate `MailMergeData` worksheet with one row per customer account.

This lets Customer Service keep working with the row-based format while Word receives a cleaner one-row-per-account mail merge source.

## Source Data Structure
The script reads the `AdjustmentData` worksheet.

Each row represents one affected service for one customer account.

| Column | Field | Description |
|---|---|---|
| A | CustomerName | Customer/account holder name |
| B | AccountNumber | MCWD account number |
| C | AddressLine1 | Mailing address |
| D | CityStateZIP | City, state, and ZIP |
| E | BillingPeriod | Example: `July 2026` or `August 2026` |
| F | Service | Service being adjusted |
| G | Difference | Adjustment amount for that service |
| H | OriginalBillTotal | Original bill total for that billing period |
| I | CorrectedBillTotal | Corrected bill total for that billing period |

Example:

```text
John Doe | 123456-789 | ... | July 2026   | Water | 2.61
John Doe | 123456-789 | ... | July 2026   | Sewer | 0.44
John Doe | 123456-789 | ... | August 2026 | Water | 13.30
John Doe | 123456-789 | ... | August 2026 | Sewer | 2.48
```

## What the Script Does
1. Opens the Excel workbook.
2. Reads the `AdjustmentData` worksheet.
3. Finds the actual last data row using the Account Number column.
4. Loads the source data into memory.
5. Uses the account number as the grouping key.
6. Groups all service rows belonging to the same account.
7. Separates July and August adjustments.
8. Collects the affected services for each billing period.
9. Collects the corresponding adjustment amounts.
10. Calculates July, August, and total account adjustments.
11. Writes one row per customer account to `MailMergeData`.
12. Saves the workbook.

## Zero-Dollar Adjustment Handling
Version 6 excludes service rows when the adjustment amount is `$0.00`.

The script rounds the adjustment amount to two decimal places before checking it:

```powershell
$difference = [math]::Round((To-Number $data.GetValue($i, $colLB + 6)), 2)

if ($difference -eq 0.0) {
    continue
}
```

### Result
If a service has a `$0.00` adjustment:

- The service is not added to the mail merge service list.
- The `$0.00` amount is not shown in the mail merge data.
- The service does not affect the July or August total.
- If all service rows for a month are `$0.00`, that month is not treated as an affected month.
- If all service rows for an account are `$0.00`, that account is not created in `MailMergeData`.

Nonzero positive or negative adjustments are still included.

## Grouping Logic
The script uses the account number as the key.

Conceptually:

```powershell
if (-not $accounts.Contains($account)) {
    $accounts[$account] = @{
        JulyServices      = @()
        JulyDifferences   = @()
        AugustServices    = @()
        AugustDifferences = @()
    }
}
```

Each additional service for the same account is added to the appropriate list.

## Building the Mail Merge Service Lines
After the rows are grouped, the script combines the service names and amounts into line-by-line values for Word.

```powershell
$julyServiceLines = [string]::Join("`n", $rec.JulyServices)

$julyDifferenceLines = [string]::Join("`n", @(
    $rec.JulyDifferences | ForEach-Object {
        '$' + ('{0:N2}' -f $_)
    }
))
```

This allows the Word letter to display a variable number of service lines depending on the customer's account.

## MailMergeData Output
The `MailMergeData` worksheet contains one row per customer account.

| Field | Description |
|---|---|
| CustomerName | Customer name |
| AccountNumber | Account number |
| AddressLine1 | Mailing address |
| CityStateZIP | City, state, ZIP |
| AffectedBillingPeriods | July, August, or both |
| JulyOriginalBill | Original July bill |
| JulyCorrectedBill | Corrected July bill |
| AugustOriginalBill | Original August bill |
| AugustCorrectedBill | Corrected August bill |
| JulyServiceLines | July affected services |
| JulyDifferenceLines | July adjustment amounts |
| JulyTotalIncrease | Total July adjustment |
| AugustServiceLines | August affected services |
| AugustDifferenceLines | August adjustment amounts |
| AugustTotalIncrease | Total August adjustment |
| TotalAdjustment | Total account adjustment |

Word should use the `MailMergeData` worksheet as the mail merge source.

## How to Run
Keep these files in the same folder:

```text
Refresh-MailMerge-Data-v6-Exclude-Zero.ps1
Refresh-MailMerge-Data-v6-Exclude-Zero.bat
2026-0918-Customer-Billing-Adjustment-Row-Based-Static-Mail-Merge-Data-TE.xlsx
```

Close the Excel workbook before running the script.

### Option 1 - Run the BAT file
Double-click:

```text
Refresh-MailMerge-Data-v6-Exclude-Zero.bat
```

### Option 2 - Run PowerShell directly

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Refresh-MailMerge-Data-v6-Exclude-Zero.ps1"
```

## Word Mail Merge
After the script finishes:

1. Open the Word customer letter.
2. Go to **Mailings**.
3. Select **Select Recipients**.
4. Choose **Use an Existing List**.
5. Select the billing adjustment Excel workbook.
6. Select the `MailMergeData` worksheet.
7. Use **Preview Results** to verify the customer information and service breakdown.

## Important Notes
- Enter source records only in `AdjustmentData`.
- Each affected service should remain on its own row.
- Multiple rows with the same account number are expected.
- The script performs the grouping automatically.
- Zero-dollar adjustments are excluded.
- Word receives one record per customer account.
- The number of service lines can vary by customer.
