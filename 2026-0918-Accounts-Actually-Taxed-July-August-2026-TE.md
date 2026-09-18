# Accounts Actually Taxed — July and August 2026

## Purpose

This query identifies customer accounts that were actually charged tax during the affected July and August 2026 billing period.

It is intended to help verify which accounts had tax applied during billing and which tax code was used.

## Criteria

The query includes records that meet all of the following conditions:

1. Transaction date is between **07/01/2026 and 08/31/2026**.
2. Transaction type is **Billing**.
3. Tax code is not NULL, blank, or whitespace.
4. Tax amount is not zero.
5. Customer account number is formatted as **999999-999**.
6. Billing cycle is included from `ub_master`.

## Tables Used

### `Springbrook0.dbo.ub_bill_detail`

Used to retrieve:

- Customer number
- Customer sequence
- Transaction date
- Transaction type
- Tax code
- Tax amount

### `Springbrook0.dbo.ub_master`

Used to retrieve:

- Billing cycle

The tables are joined using:

- `cust_no`
- `cust_sequence`

## Output Columns

| Column | Description |
|---|---|
| `account_number` | Customer account number formatted as `999999-999` |
| `billing_cycle` | Billing cycle from `ub_master` |
| `transaction_date` | Billing transaction date formatted as `MM/DD/YYYY` |
| `tran_type` | Transaction type; limited to `Billing` |
| `tax_code` | Tax code actually recorded on the billing transaction |
| `amount` | Tax amount recorded for the transaction |

## SQL Query

```sql
/* ============================================================
   Accounts Actually Taxed
   Criteria:
   1. Transaction date between 07/01/2026 and 08/31/2026
   2. Transaction type = Billing
   3. Tax code is NOT NULL or blank
   4. Tax amount is not zero
   5. Account formatted as 999999-999
   6. Includes billing cycle from ub_master
   ============================================================ */

SELECT
      RIGHT('000000' + CAST(BD.[cust_no] AS varchar(6)), 6)
        + '-'
        + RIGHT('000' + CAST(BD.[cust_sequence] AS varchar(3)), 3)
        AS [account_number]

    , M.[billing_cycle]

    , CONVERT(varchar(10), BD.[tran_date], 101)
        AS [transaction_date]

    , BD.[tran_type]
    , BD.[tax_code]
    , BD.[amount]

FROM [Springbrook0].[dbo].[ub_bill_detail] BD

INNER JOIN [Springbrook0].[dbo].[ub_master] M
    ON M.[cust_no] = BD.[cust_no]
    AND M.[cust_sequence] = BD.[cust_sequence]

WHERE
    BD.[tran_date] >= '2026-07-01'
    AND BD.[tran_date] < '2026-09-01'

    AND BD.[tran_type] = 'Billing'

    AND NULLIF(LTRIM(RTRIM(BD.[tax_code])), '') IS NOT NULL

    AND BD.[amount] <> 0

ORDER BY
      M.[billing_cycle]
    , [account_number]
    , BD.[tran_date]
    , BD.[tax_code];
```

## Notes

- The date filter uses an inclusive start date of `07/01/2026` and an exclusive end date of `09/01/2026`. This captures all transactions dated in July and August 2026.
- Records with no tax code are excluded.
- Records with a zero tax amount are excluded.
- The query reports actual billing detail records, so an account may appear more than once if it had multiple taxable billing transactions or multiple tax codes during the period.
- The report is read-only and does not make any changes to Springbrook data.
