# Inactive Accounts Billed During July–August 2026

## Purpose

This SQL report identifies customer accounts that:

1. **Were billed between July 1, 2026 and August 31, 2026.**
2. **Were still active when the billing transaction occurred.**
3. **Have since become inactive/finalized.**

The report is intended to identify accounts that were part of the affected July/August 2026 billing period but should be reviewed separately because they are no longer active.

## How the Query Works

The query starts with the `ub_master` table and returns accounts that have a `final_date` in the past.

It then checks the `ub_bill_detail` table to confirm that the account had at least one billing transaction during the affected period.

An account is included only when all of the following conditions are true:

- `final_date` is not `NULL`.
- `final_date` is earlier than the current date.
- A billing transaction exists between **07/01/2026** and **08/31/2026**.
- The account's `final_date` is on or after the billing transaction date, confirming the account had not yet been finalized when the transaction occurred.

## Output Fields

| Field | Description |
|---|---|
| `account_number` | Customer number and sequence formatted as `999999-999`. |
| `acct_status` | Current Springbrook account status. |
| `billing_cycle` | Billing cycle assigned to the account. |
| `connect_date` | Date the account/service was connected. |
| `final_date` | Date the account was finalized or became inactive. |
| `last_bill_date` | Most recent billing date stored in `ub_master`. |

## Important Notes

- The query uses the **actual billing transaction date** from `ub_bill_detail.tran_date` rather than relying only on `last_bill_date`.
- The date range uses:
  - `>= '2026-07-01'`
  - `< '2026-09-01'`

  This safely includes all billing transactions dated through **August 31, 2026**, regardless of whether `tran_date` contains a time value.

- `EXISTS` is used so that each account appears only once even if it had multiple billing-detail records during July and August.
- The query does not require the current `acct_status` to equal a specific value such as `Final` or `Delete`. Instead, the presence of a past `final_date` is used to determine that the account has since become inactive.
- Because the query uses `GETDATE()`, the report should be rerun before the final adjustment process so that accounts finalized after an earlier report are also captured.

## SQL

```sql
/* ============================================================
   Accounts that:
   1. Were billed between 07/01/2026 and 08/31/2026
   2. Were still active when that billing occurred
   3. Have since become inactive
   ============================================================ */

SELECT
      RIGHT('000000' + CAST(m.[cust_no] AS varchar(6)), 6)
        + '-'
        + RIGHT('000' + CAST(m.[cust_sequence] AS varchar(3)), 3)
        AS [account_number]

    , m.[acct_status]
    , m.[billing_cycle]

    , CONVERT(varchar(10), m.[connect_date], 101) AS [connect_date]
    , CONVERT(varchar(10), m.[final_date], 101) AS [final_date]
    , CONVERT(varchar(10), m.[last_bill_date], 101) AS [last_bill_date]

FROM [Springbrook0].[dbo].[ub_master] m

WHERE
    -- Account has since become inactive
    m.[final_date] IS NOT NULL
    AND m.[final_date] < CAST(GETDATE() AS date)

    -- Account had an actual billing transaction during
    -- July or August 2026 while the account was still active
    AND EXISTS
    (
        SELECT 1
        FROM [Springbrook0].[dbo].[ub_bill_detail] b

        WHERE
            b.[cust_no] = m.[cust_no]
            AND b.[cust_sequence] = m.[cust_sequence]

            -- July 1 through August 31, 2026
            AND b.[tran_date] >= '2026-07-01'
            AND b.[tran_date] < '2026-09-01'

            -- Account had not yet been finalized
            -- when this billing transaction occurred
            AND m.[final_date] >= CAST(b.[tran_date] AS date)
    )

ORDER BY
    m.[billing_cycle],
    CONVERT(varchar(10), m.[final_date], 101) DESC,
    [account_number];
```

## Expected Use

This report can be used as part of the July/August 2026 billing-adjustment review to identify accounts that were valid and billable during the affected period but are no longer active today.

These accounts can then be reviewed before adjustment imports or customer notices are finalized so inactive accounts are handled appropriately.
