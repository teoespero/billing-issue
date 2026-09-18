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