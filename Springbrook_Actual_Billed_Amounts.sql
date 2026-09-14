/*==========================================================================
    SCRIPT NAME:
        Springbrook_Actual_Billed_Amounts.sql

    PURPOSE:
        Reports what Springbrook ACTUALLY billed.

        No rate recalculation.
        No revision recalculation.
        No adjustment calculation.

        Actual billed amount:
            dbo.ub_bill_detail.amount

        Includes:
            - All billing cycles
            - All service numbers found in dbo.ub_service
            - Account
            - Billing period
            - Service
            - Detail code
            - Consumption
            - Billable consumption
            - Revision used
            - Actual amount billed

    DATABASE:
        Springbrook0
==========================================================================*/


/*==========================================================================
    PARAMETERS

    End date is exclusive.

    Example:
        07/01/2026 through 08/31/2026

        @StartDate = '2026-07-01'
        @EndDate   = '2026-09-01'
==========================================================================*/

DECLARE @StartDate DATE = '2026-07-01';
DECLARE @EndDate   DATE = '2026-09-01';


/*==========================================================================
    ACTUAL BILLED DETAIL
==========================================================================*/

SELECT

    /*----------------------------------------------------------------------
        ACCOUNT
    ----------------------------------------------------------------------*/

    RIGHT
    (
        '000000' + CAST(bd.cust_no AS VARCHAR(6)),
        6
    )
    + '-'
    +
    RIGHT
    (
        '000' + CAST(bd.cust_sequence AS VARCHAR(3)),
        3
    )                                           AS [Account No],


    m.billing_cycle                             AS [Billing Cycle],

    m.lot_no                                    AS [Lot No],

    l.misc_1                                    AS [Area],

    l.misc_2                                    AS [ST_Category],


    /*----------------------------------------------------------------------
        BILL / TRANSACTION
    ----------------------------------------------------------------------*/

    bd.transaction_id                           AS [Transaction ID],

    CONVERT
    (
        VARCHAR(10),
        bd.tran_date,
        101
    )                                           AS [Transaction Date],


    CONVERT
    (
        VARCHAR(10),
        bd.period_begin_date,
        101
    )                                           AS [Period Begin],


    CONVERT
    (
        VARCHAR(10),
        bd.period_end_date,
        101
    )                                           AS [Period End],


    CASE
        WHEN bd.period_begin_date IS NOT NULL
         AND bd.period_end_date IS NOT NULL
        THEN
            DATEDIFF
            (
                DAY,
                bd.period_begin_date,
                bd.period_end_date
            ) + 1
    END                                         AS [Billing Days],


    /*----------------------------------------------------------------------
        SERVICE

        service_number comes directly from dbo.ub_service
    ----------------------------------------------------------------------*/

    s.service_number                            AS [Service Number],

    s.service_code                              AS [Service Code],

    s.description                               AS [Service Description],

    s.bill_type                                 AS [Bill Type],


    /*----------------------------------------------------------------------
        BILL DETAIL
    ----------------------------------------------------------------------*/

    bd.code                                     AS [Detail Code],

    bd.revision_no                              AS [Revision Used],

    bd.use_period                               AS [Rate Period],

    bd.consumption                              AS [Consumption],

    bd.billable_cons                            AS [Billable Consumption],

    bd.pcnt_of_period                           AS [Percent of Period],


    /*======================================================================
        ACTUAL AMOUNT BILLED BY SPRINGBROOK
    ======================================================================*/

    CAST
    (
        bd.amount
        AS DECIMAL(18,2)
    )                                           AS [Amount Billed],


    /*======================================================================
        TOTAL BILL

        This gives the total of all service/detail lines belonging to
        the same customer billing transaction.
    ======================================================================*/

    CAST
    (
        SUM(bd.amount) OVER
        (
            PARTITION BY
                bd.cust_no,
                bd.cust_sequence,
                bd.transaction_id
        )
        AS DECIMAL(18,2)
    )                                           AS [Total Bill Amount]


FROM [Springbrook0].[dbo].[ub_bill_detail] AS bd


/*==========================================================================
    ACCOUNT / BILLING CYCLE
==========================================================================*/

INNER JOIN [Springbrook0].[dbo].[ub_master] AS m

    ON  bd.cust_no = m.cust_no
    AND bd.cust_sequence = m.cust_sequence


/*==========================================================================
    SERVICE

    This makes sure the billed service exists in dbo.ub_service.

    No service-number filtering is done, so ALL service numbers are included.
==========================================================================*/

INNER JOIN [Springbrook0].[dbo].[ub_service] AS s

    ON bd.service_code = s.service_code


/*==========================================================================
    LOT INFORMATION
==========================================================================*/

LEFT JOIN [Springbrook0].[dbo].[lot] AS l

    ON m.lot_no = l.lot_no


/*==========================================================================
    FILTERS
==========================================================================*/

WHERE

    bd.tran_date >= @StartDate

    AND bd.tran_date < @EndDate

    AND bd.tran_type = 'BILLING'


/*==========================================================================
    SORT
==========================================================================*/

ORDER BY

    m.billing_cycle,

    bd.cust_no,
    bd.cust_sequence,

    bd.tran_date,

    s.service_number,

    bd.code;