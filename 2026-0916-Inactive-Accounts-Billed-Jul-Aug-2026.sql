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
	CONVERT(varchar(10), m.[final_date], 101) desc,
    [account_number];