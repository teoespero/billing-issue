/*==========================================================================
    SCRIPT NAME:
        Springbrook_Billing_Adjustment_Calculation.sql

    PURPOSE:
        Recalculates Springbrook service charges after customers have
        already been billed and determines the required adjustment.

        IMPORTANT:

        Springbrook may store ONE logical billing line as MULTIPLE
        ub_bill_detail records when the billing period crosses rate
        revisions.

        Example:

            WA011 Consumption

                Rev 22 = $70.28
                Rev 23 = $18.75

            These are NOT two separate consumption charges.

            Original Logical Line Total:

                $70.28 + $18.75 = $89.03


        Therefore this script FIRST collapses Springbrook's original
        revision rows into ONE logical billing line.

        The corrected rate calculation is then performed ONCE against
        that logical billing line.


    ========================================================================
    ADJUSTMENT FORMULA
    ========================================================================

        Correct Amount - Original Billed Amount = Adjustment

        Positive = CHARGE
        Negative = CREDIT
        Zero     = NO ADJUSTMENT


    ========================================================================
    ORIGINAL BILLING LINE CONSOLIDATION
    ========================================================================

        Original Springbrook rows are grouped by:

            Customer
            Customer Sequence
            Transaction
            Service Code
            Detail Code
            Rate Period
            Period Begin
            Period End

        Original Amount:

            SUM(bd.amount)

        Billable Consumption:

            MAX(bd.billable_cons)

        Number of Units:

            MAX(bd.no_of_units)

        Percent of Period:

            SUM(bd.pcnt_of_period)


    ========================================================================
    CONSUMPTION RATE CALCULATION
    ========================================================================

        1. Use Springbrook's original billable_cons.

        2. Determine FULL-BILL consumption levels FIRST.

        3. Determine all corrected service-rate revisions effective
           during the billing period.

        4. Determine billing days belonging to each revision.

        5. Prorate the established full-bill tier units between revisions.

        6. Apply the SAME consumption level from each applicable revision.

        7. Sum all revision / tier components.


    ========================================================================
    FLAT / MINIMUM CALCULATION
    ========================================================================

        Revision Minimum
            x
        Number of Units
            x
        Revision Days / Total Billing Days


    ========================================================================
    OUTPUTS
    ========================================================================

        RESULT SET 1:
            Pivoted Billing Adjustment Summary

            ONE ROW PER:
                Account + Billing Transaction


        RESULT SET 2:
            Billing Line Adjustment Detail

            ONE ROW PER LOGICAL BILLING LINE.

            Example:

                SW01 Flat
                WA011 Consumption
                WA011 Flat


    DATABASE:
        Springbrook0


    AUTHOR:
        Teo Espero


    REVISION HISTORY:
    ------------------------------------------------------------------------
    Version     Date          Author        Description
    ------------------------------------------------------------------------
    1.0         09/04/2026    Teo Espero    Initial calculation.

    2.0         09/14/2026    Teo Espero    Added revision proration.

    3.0         09/14/2026    Teo Espero    Expanded to all qualifying
                                             service rates.

    3.1         09/14/2026    Teo Espero    Added dynamic pivot.

    3.2         09/14/2026    Teo Espero    Added No. of Units multiplier.

    3.3         09/14/2026    Teo Espero    Uses Springbrook billable_cons
                                             as authoritative consumption.

    4.0         09/14/2026    Teo Espero    Collapses original Springbrook
                                             revision rows BEFORE rate
                                             recalculation.

                                             Prevents duplicate logical
                                             billing lines and duplicate
                                             adjustments.

                                             Corrected WaterUsage duplicate
                                             revision counting.

==========================================================================*/


/*==========================================================================
    PARAMETERS
==========================================================================*/

DECLARE @BillingCycle INT  = 1;
DECLARE @StartDate    DATE = '2026-07-01';
DECLARE @EndDate      DATE = '2026-08-01';



/*==========================================================================
    CLEANUP FROM PRIOR EXECUTION
==========================================================================*/

DROP TABLE IF EXISTS #WaterUsage;
DROP TABLE IF EXISTS #RawBillLines;
DROP TABLE IF EXISTS #BillLines;
DROP TABLE IF EXISTS #ServiceRevisionWindows;
DROP TABLE IF EXISTS #ConsumptionLevels;
DROP TABLE IF EXISTS #RevisionSegmentCalc;
DROP TABLE IF EXISTS #ConsumptionComponents;
DROP TABLE IF EXISTS #FlatComponents;
DROP TABLE IF EXISTS #CorrectedConsumption;
DROP TABLE IF EXISTS #CorrectedFlat;
DROP TABLE IF EXISTS #RevisionSummary;
DROP TABLE IF EXISTS #FinalLines;
DROP TABLE IF EXISTS #TransactionSummary;
DROP TABLE IF EXISTS #PivotSource;



/*==========================================================================
    1. WATER BILLABLE CONSUMPTION - AUDIT ONLY

    IMPORTANT:

        Springbrook may have multiple Water Consumption rows because the
        original bill crossed rate revisions.

        DO NOT SUM billable_cons directly from ub_bill_detail.

        Example:

            Rev 22 billable_cons = 14
            Rev 23 billable_cons = 14

        Actual consumption = 14, NOT 28.

        First collapse each logical Water Consumption line using MAX(),
        then total logical Water lines if necessary.
==========================================================================*/

;WITH WaterLogicalLines AS
(
    SELECT
        bd.cust_no,
        bd.cust_sequence,
        bd.transaction_id,

        bd.service_code,

        bd.period_begin_date,
        bd.period_end_date,


        CAST
        (
            MAX
            (
                COALESCE
                (
                    bd.billable_cons,
                    bd.consumption,
                    0
                )
            )

            AS DECIMAL(18,6)

        ) AS water_billable_cons


    FROM [Springbrook0].[dbo].[ub_bill_detail] AS bd


    INNER JOIN [Springbrook0].[dbo].[ub_master] AS m

        ON  bd.cust_no = m.cust_no
        AND bd.cust_sequence = m.cust_sequence


    WHERE
        m.billing_cycle = @BillingCycle

        AND bd.tran_date >= @StartDate
        AND bd.tran_date <  @EndDate

        AND bd.tran_type = 'BILLING'

        AND bd.service_code LIKE 'WA%'

        AND UPPER(bd.code) = 'CONSUMPTION'


    GROUP BY
        bd.cust_no,
        bd.cust_sequence,
        bd.transaction_id,

        bd.service_code,

        bd.period_begin_date,
        bd.period_end_date
)


SELECT
    cust_no,
    cust_sequence,
    transaction_id,


    CAST
    (
        SUM(water_billable_cons)

        AS DECIMAL(18,6)

    ) AS water_billable_cons


INTO #WaterUsage


FROM WaterLogicalLines


GROUP BY
    cust_no,
    cust_sequence,
    transaction_id;



/*==========================================================================
    2. GET RAW SPRINGBROOK BILL DETAIL ROWS

    DO NOT RECALCULATE DIRECTLY FROM THIS TABLE.

    These rows will be consolidated into logical billing lines in Step 3.
==========================================================================*/

SELECT
    bd.ub_bill_detail_id,

    bd.transaction_id,

    bd.cust_no,
    bd.cust_sequence,


    m.billing_cycle,

    m.lot_no,


    l.misc_1 AS area,

    l.misc_2 AS st_category,


    bd.tran_date,

    bd.period_begin_date,
    bd.period_end_date,


    bd.service_code,

    bd.code AS detail_code,


    bd.revision_no
        AS billed_revision,


    /*----------------------------------------------------------------------
        RATE PERIOD
    ----------------------------------------------------------------------*/

    COALESCE
    (
        NULLIF
        (
            bd.use_period,
            0
        ),

        1

    ) AS use_period,


    /*----------------------------------------------------------------------
        NUMBER OF UNITS
    ----------------------------------------------------------------------*/

    CAST
    (
        COALESCE
        (
            NULLIF
            (
                bd.no_of_units,
                0
            ),

            1
        )

        AS DECIMAL(18,6)

    ) AS no_of_units,


    /*----------------------------------------------------------------------
        ORIGINAL CONSUMPTION
    ----------------------------------------------------------------------*/

    CAST
    (
        bd.consumption

        AS DECIMAL(18,6)

    ) AS consumption,


    /*----------------------------------------------------------------------
        ORIGINAL BILLABLE CONSUMPTION
    ----------------------------------------------------------------------*/

    CAST
    (
        bd.billable_cons

        AS DECIMAL(18,6)

    ) AS billable_cons,


    /*----------------------------------------------------------------------
        ORIGINAL PERCENT OF PERIOD

        Example:

            Rev 22 = 81.81
            Rev 23 = 18.18
    ----------------------------------------------------------------------*/

    bd.pcnt_of_period,


    /*----------------------------------------------------------------------
        ORIGINAL AMOUNT FOR THIS SPRINGBROOK REVISION ROW
    ----------------------------------------------------------------------*/

    CAST
    (
        bd.amount

        AS DECIMAL(18,2)

    ) AS springbrook_amount,


    /*----------------------------------------------------------------------
        SERVICE INFORMATION
    ----------------------------------------------------------------------*/

    s.ub_service_id,

    s.service_number,

    s.description
        AS service_description


INTO #RawBillLines


FROM [Springbrook0].[dbo].[ub_bill_detail] AS bd


INNER JOIN [Springbrook0].[dbo].[ub_master] AS m

    ON  bd.cust_no = m.cust_no
    AND bd.cust_sequence = m.cust_sequence


INNER JOIN [Springbrook0].[dbo].[ub_service] AS s

    ON bd.service_code = s.service_code


LEFT JOIN [Springbrook0].[dbo].[lot] AS l

    ON m.lot_no = l.lot_no


WHERE
    m.billing_cycle = @BillingCycle

    AND bd.tran_date >= @StartDate
    AND bd.tran_date <  @EndDate

    AND bd.tran_type = 'BILLING'


    /*----------------------------------------------------------------------
        CURRENTLY SUPPORTED BILLING CALCULATION TYPES
    ----------------------------------------------------------------------*/

    AND UPPER(bd.code) IN
    (
        'CONSUMPTION',
        'FLAT'
    )


    AND bd.period_begin_date IS NOT NULL

    AND bd.period_end_date IS NOT NULL

    AND bd.period_end_date
        >= bd.period_begin_date;



/*==========================================================================
    3. COLLAPSE ORIGINAL SPRINGBROOK REVISION ROWS

    THIS IS THE IMPORTANT DUPLICATE FIX.

    Example original records:

        WA011 Consumption

            Rev 22     81.81%     $70.28
            Rev 23     18.18%     $18.75


    Result:

        ONE logical WA011 Consumption line

            Original Amount = $89.03

==========================================================================*/

;WITH LogicalBillLines AS
(
    SELECT
        rbl.cust_no,

        rbl.cust_sequence,

        rbl.transaction_id,


        MAX(rbl.billing_cycle)
            AS billing_cycle,


        MAX(rbl.lot_no)
            AS lot_no,


        MAX(rbl.area)
            AS area,


        MAX(rbl.st_category)
            AS st_category,


        MAX(rbl.tran_date)
            AS tran_date,


        rbl.period_begin_date,

        rbl.period_end_date,


        rbl.service_code,

        rbl.detail_code,

        rbl.use_period,


        MAX(rbl.ub_service_id)
            AS ub_service_id,


        MAX(rbl.service_number)
            AS service_number,


        MAX(rbl.service_description)
            AS service_description,


        /*------------------------------------------------------------------
            NUMBER OF RAW SPRINGBROOK DETAIL ROWS COLLAPSED
        ------------------------------------------------------------------*/

        COUNT(*)
            AS source_detail_rows,


        /*------------------------------------------------------------------
            FIRST SOURCE DETAIL ID - AUDIT ONLY
        ------------------------------------------------------------------*/

        MIN(rbl.ub_bill_detail_id)
            AS first_source_detail_id,


        /*------------------------------------------------------------------
            NUMBER OF UNITS

            Revision-split rows should carry the same No. of Units.
        ------------------------------------------------------------------*/

        CAST
        (
            MAX(rbl.no_of_units)

            AS DECIMAL(18,6)

        ) AS no_of_units,


        /*------------------------------------------------------------------
            CONSUMPTION

            Do NOT sum duplicate revision rows.
        ------------------------------------------------------------------*/

        CAST
        (
            MAX(rbl.consumption)

            AS DECIMAL(18,6)

        ) AS consumption,


        /*------------------------------------------------------------------
            BILLABLE CONSUMPTION

            Do NOT sum duplicate revision rows.

            Example:
                14 + 14 does NOT equal 28.

                The logical consumption is 14.
        ------------------------------------------------------------------*/

        CAST
        (
            MAX(rbl.billable_cons)

            AS DECIMAL(18,6)

        ) AS original_billable_cons,


        /*------------------------------------------------------------------
            TOTAL ORIGINAL PERCENT

            Generally ~100 after combining revisions.
        ------------------------------------------------------------------*/

        CAST
        (
            SUM
            (
                COALESCE
                (
                    rbl.pcnt_of_period,
                    0
                )
            )

            AS DECIMAL(18,4)

        ) AS original_percent_of_period,


        /*------------------------------------------------------------------
            ORIGINAL BILLED AMOUNT

            THIS IS THE TRUE LOGICAL LINE AMOUNT.

            Example:

                $70.28 + $18.75 = $89.03
        ------------------------------------------------------------------*/

        CAST
        (
            SUM
            (
                COALESCE
                (
                    rbl.springbrook_amount,
                    0
                )
            )

            AS DECIMAL(18,2)

        ) AS springbrook_amount,


        /*------------------------------------------------------------------
            ORIGINAL BILLED REVISION LIST
        ------------------------------------------------------------------*/

        STRING_AGG
        (
            CAST
            (
                'Rev '
                + CAST
                  (
                      rbl.billed_revision
                      AS VARCHAR(20)
                  )

                AS VARCHAR(MAX)
            ),

            ' | '

        ) WITHIN GROUP
        (
            ORDER BY
                rbl.billed_revision,
                rbl.ub_bill_detail_id
        ) AS billed_revisions,


        /*------------------------------------------------------------------
            ORIGINAL SPRINGBROOK BILL BREAKDOWN

            Example:

                Rev 22 81.81% = $70.28
                |
                Rev 23 18.18% = $18.75
        ------------------------------------------------------------------*/

        STRING_AGG
        (
            CAST
            (
                'Rev '

                + CAST
                  (
                      rbl.billed_revision
                      AS VARCHAR(20)
                  )

                + ' '

                + CAST
                  (
                      CAST
                      (
                          COALESCE
                          (
                              rbl.pcnt_of_period,
                              0
                          )

                          AS DECIMAL(10,2)
                      )

                      AS VARCHAR(20)
                  )

                + '% = $'

                + CAST
                  (
                      CAST
                      (
                          rbl.springbrook_amount

                          AS DECIMAL(18,2)
                      )

                      AS VARCHAR(30)
                  )

                AS VARCHAR(MAX)
            ),

            ' | '

        ) WITHIN GROUP
        (
            ORDER BY
                rbl.billed_revision,
                rbl.ub_bill_detail_id
        ) AS original_billed_breakdown


    FROM #RawBillLines AS rbl


    GROUP BY
        rbl.cust_no,

        rbl.cust_sequence,

        rbl.transaction_id,

        rbl.service_code,

        rbl.detail_code,

        rbl.use_period,

        rbl.period_begin_date,

        rbl.period_end_date
)


SELECT

    /*----------------------------------------------------------------------
        NEW LOGICAL BILL LINE ID

        This replaces ub_bill_detail_id for recalculation purposes.
    ----------------------------------------------------------------------*/

    ROW_NUMBER() OVER
    (
        ORDER BY
            lbl.cust_no,
            lbl.cust_sequence,
            lbl.transaction_id,
            lbl.service_code,
            lbl.detail_code,
            lbl.period_begin_date,
            lbl.period_end_date,
            lbl.use_period

    ) AS logical_bill_line_id,


    lbl.*,


    /*----------------------------------------------------------------------
        WATER CONSUMPTION - AUDIT ONLY
    ----------------------------------------------------------------------*/

    wu.water_billable_cons,


    /*----------------------------------------------------------------------
        WINTER AVERAGE - AUDIT ONLY
    ----------------------------------------------------------------------*/

    wa.winter_average,

    wa.winter_average_effective_date,


    /*======================================================================
        BILLABLE CONSUMPTION USED FOR RECALCULATION

        The original Springbrook logical Billable Consumption is retained.
    ======================================================================*/

    lbl.original_billable_cons
        AS billable_cons,


    CASE

        WHEN UPPER(lbl.detail_code) = 'CONSUMPTION'

            THEN 'SPRINGBROOK BILLABLE CONS'

        ELSE 'NOT APPLICABLE'

    END AS consumption_basis


INTO #BillLines


FROM LogicalBillLines AS lbl


LEFT JOIN #WaterUsage AS wu

    ON  lbl.cust_no = wu.cust_no
    AND lbl.cust_sequence = wu.cust_sequence
    AND lbl.transaction_id = wu.transaction_id



/*--------------------------------------------------------------------------
    WINTER AVERAGE - AUDIT ONLY
--------------------------------------------------------------------------*/

OUTER APPLY
(
    SELECT TOP (1)

        CAST
        (
            w.winter_average
            AS DECIMAL(18,6)

        ) AS winter_average,


        w.effective_date
            AS winter_average_effective_date


    FROM [Springbrook0].[dbo].[ub_winter_average] AS w


    WHERE
        w.cust_no = lbl.cust_no

        AND w.cust_sequence = lbl.cust_sequence

        AND w.service_number = lbl.service_number

        AND w.record_committed = 1

        AND w.effective_date
            <= lbl.period_end_date


    ORDER BY
        w.effective_date DESC,
        w.ub_winter_average_id DESC

) AS wa;



/*==========================================================================
    4. BUILD CLEAN SERVICE RATE REVISION LIST
==========================================================================*/

;WITH ServiceRevisionRank AS
(
    SELECT
        sd.ub_service_detail_id,

        sd.ub_service_id,

        sd.revision_no,

        sd.effective_date,

        sd.minimum,


        ROW_NUMBER() OVER
        (
            PARTITION BY
                sd.ub_service_id,
                sd.effective_date

            ORDER BY
                sd.revision_no DESC,
                sd.ub_service_detail_id DESC

        ) AS same_date_rank


    FROM [Springbrook0].[dbo].[ub_service_detail] AS sd


    WHERE
        sd.effective_date IS NOT NULL
),

CleanRevision AS
(
    SELECT
        ub_service_detail_id,

        ub_service_id,

        revision_no,

        effective_date,

        minimum


    FROM ServiceRevisionRank


    WHERE
        same_date_rank = 1
)


SELECT
    ub_service_detail_id,

    ub_service_id,

    revision_no,

    effective_date,

    minimum,


    LEAD
    (
        effective_date

    ) OVER
    (
        PARTITION BY
            ub_service_id

        ORDER BY
            effective_date,
            revision_no,
            ub_service_detail_id

    ) AS next_effective_date


INTO #ServiceRevisionWindows


FROM CleanRevision;



/*==========================================================================
    5. BUILD CONSUMPTION LEVEL DEFINITIONS
==========================================================================*/

SELECT
    scl.ub_service_cons_lvl_id,

    scl.ub_service_detail_id,

    scl.period_number,

    scl.cons_level,


    COALESCE
    (
        scl.minimum_cons,
        0

    ) AS minimum_cons,


    scl.rate,


    /*----------------------------------------------------------------------
        NEXT LEVEL START
    ----------------------------------------------------------------------*/

    LEAD
    (
        COALESCE
        (
            scl.minimum_cons,
            0
        )

    ) OVER
    (
        PARTITION BY
            scl.ub_service_detail_id,
            scl.period_number

        ORDER BY
            scl.cons_level,
            scl.ub_service_cons_lvl_id

    ) AS next_minimum_cons,


    /*----------------------------------------------------------------------
        NUMBER OF LEVELS
    ----------------------------------------------------------------------*/

    COUNT(*) OVER
    (
        PARTITION BY
            scl.ub_service_detail_id,
            scl.period_number

    ) AS level_count


INTO #ConsumptionLevels


FROM [Springbrook0].[dbo].[ub_service_cons_lvl] AS scl;



/*==========================================================================
    6. FIND CORRECT RATE REVISIONS OVERLAPPING BILLING PERIOD

    NOTE:

        This occurs AFTER original Springbrook revision rows have already
        been collapsed into one logical line.
==========================================================================*/

;WITH RevisionSegments AS
(
    SELECT
        bl.*,


        srw.ub_service_detail_id
            AS revision_service_detail_id,


        srw.revision_no
            AS applicable_revision,


        srw.effective_date
            AS revision_effective_date,


        srw.next_effective_date,


        srw.minimum
            AS revision_minimum,


        /*------------------------------------------------------------------
            FIRST DAY USING REVISION
        ------------------------------------------------------------------*/

        CASE

            WHEN srw.effective_date
                 > bl.period_begin_date

                THEN srw.effective_date

            ELSE bl.period_begin_date

        END AS revision_period_begin,


        /*------------------------------------------------------------------
            LAST DAY USING REVISION
        ------------------------------------------------------------------*/

        CASE

            WHEN
                srw.next_effective_date IS NOT NULL

                AND srw.next_effective_date
                    <= bl.period_end_date

                THEN
                    DATEADD
                    (
                        DAY,
                        -1,
                        srw.next_effective_date
                    )

            ELSE bl.period_end_date

        END AS revision_period_end


    FROM #BillLines AS bl


    INNER JOIN #ServiceRevisionWindows AS srw

        ON bl.ub_service_id
           = srw.ub_service_id


        AND srw.effective_date
            <= bl.period_end_date


        AND
        (
            srw.next_effective_date IS NULL

            OR

            srw.next_effective_date
                > bl.period_begin_date
        )
)


SELECT
    rs.*,


    /*----------------------------------------------------------------------
        TOTAL BILLING DAYS
    ----------------------------------------------------------------------*/

    DATEDIFF
    (
        DAY,
        rs.period_begin_date,
        rs.period_end_date

    ) + 1 AS total_billing_days,


    /*----------------------------------------------------------------------
        DAYS USING THIS REVISION
    ----------------------------------------------------------------------*/

    DATEDIFF
    (
        DAY,
        rs.revision_period_begin,
        rs.revision_period_end

    ) + 1 AS revision_days,


    /*----------------------------------------------------------------------
        REVISION RATIO
    ----------------------------------------------------------------------*/

    CAST
    (
        (
            DATEDIFF
            (
                DAY,
                rs.revision_period_begin,
                rs.revision_period_end

            ) + 1
        )

        * 1.0

        /

        NULLIF
        (
            DATEDIFF
            (
                DAY,
                rs.period_begin_date,
                rs.period_end_date

            ) + 1,

            0
        )

        AS DECIMAL(18,10)

    ) AS revision_ratio


INTO #RevisionSegmentCalc


FROM RevisionSegments AS rs


WHERE
    rs.revision_period_end
        >= rs.revision_period_begin;



/*==========================================================================
    7. DETERMINE FULL-BILL CONSUMPTION LEVELS

    FULL BILLABLE CONSUMPTION determines tiers FIRST.

    Tier units are then allocated between corrected rate revisions.
==========================================================================*/

;WITH FullBillConsumptionLevels AS
(
    SELECT
        rsc.*,

        cl.period_number,

        cl.cons_level,

        cl.level_count,

        cl.minimum_cons,

        cl.next_minimum_cons,

        cl.rate,


        CAST
        (
            CASE

                /*----------------------------------------------------------
                    ONLY ONE LEVEL
                ----------------------------------------------------------*/

                WHEN cl.level_count = 1

                    THEN rsc.billable_cons


                /*----------------------------------------------------------
                    DID NOT REACH LEVEL
                ----------------------------------------------------------*/

                WHEN rsc.billable_cons
                     <= cl.minimum_cons

                    THEN 0


                /*----------------------------------------------------------
                    LAST / OPEN LEVEL
                ----------------------------------------------------------*/

                WHEN cl.next_minimum_cons IS NULL

                    THEN
                        rsc.billable_cons
                        -
                        cl.minimum_cons


                /*----------------------------------------------------------
                    FULLY FILLED LEVEL
                ----------------------------------------------------------*/

                WHEN rsc.billable_cons
                     >= cl.next_minimum_cons

                    THEN
                        cl.next_minimum_cons
                        -
                        cl.minimum_cons


                /*----------------------------------------------------------
                    PARTIALLY FILLED LEVEL
                ----------------------------------------------------------*/

                ELSE
                    rsc.billable_cons
                    -
                    cl.minimum_cons

            END

            AS DECIMAL(18,6)

        ) AS full_bill_level_units


    FROM #RevisionSegmentCalc AS rsc


    INNER JOIN #ConsumptionLevels AS cl

        ON rsc.revision_service_detail_id
           = cl.ub_service_detail_id


        AND cl.period_number
            = rsc.use_period


    WHERE
        UPPER(rsc.detail_code)
        = 'CONSUMPTION'
),


ProratedConsumptionLevels AS
(
    SELECT
        fbcl.*,


        CAST
        (
            fbcl.full_bill_level_units
            *
            fbcl.revision_ratio

            AS DECIMAL(18,6)

        ) AS revision_level_units


    FROM FullBillConsumptionLevels AS fbcl


    WHERE
        fbcl.full_bill_level_units > 0
)


SELECT
    pcl.*,


    CAST
    (
        pcl.revision_level_units
        *
        pcl.rate

        AS DECIMAL(18,2)

    ) AS component_amount


INTO #ConsumptionComponents


FROM ProratedConsumptionLevels AS pcl


WHERE
    pcl.revision_level_units > 0;



/*==========================================================================
    8. FLAT / MINIMUM COMPONENTS

    Flat Component:

        Revision Minimum
            x
        Number of Units
            x
        Revision Ratio
==========================================================================*/

SELECT
    rsc.*,


    CAST
    (
        rsc.revision_minimum
        *
        rsc.no_of_units
        *
        rsc.revision_ratio

        AS DECIMAL(18,2)

    ) AS component_amount


INTO #FlatComponents


FROM #RevisionSegmentCalc AS rsc


WHERE
    UPPER(rsc.detail_code)
    = 'FLAT';



/*==========================================================================
    9. AGGREGATE CORRECTED CONSUMPTION

    ONE RESULT PER LOGICAL BILLING LINE
==========================================================================*/

SELECT
    logical_bill_line_id,


    STRING_AGG
    (
        CAST
        (
            'Rev '

            + CAST
              (
                  applicable_revision
                  AS VARCHAR(20)
              )

            + ' ['

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_begin,
                  101
              )

            + '-'

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_end,
                  101
              )

            + '] '

            + CAST
              (
                  revision_days
                  AS VARCHAR(10)
              )

            + '/'

            + CAST
              (
                  total_billing_days
                  AS VARCHAR(10)
              )

            + ' days; L'

            + CAST
              (
                  cons_level
                  AS VARCHAR(10)
              )

            + ' '

            + CAST
              (
                  CAST
                  (
                      revision_level_units
                      AS DECIMAL(18,3)
                  )

                  AS VARCHAR(30)
              )

            + ' x $'

            + CAST
              (
                  CAST
                  (
                      rate
                      AS DECIMAL(18,4)
                  )

                  AS VARCHAR(30)
              )

            + ' = $'

            + CAST
              (
                  CAST
                  (
                      component_amount
                      AS DECIMAL(18,2)
                  )

                  AS VARCHAR(30)
              )

            AS VARCHAR(MAX)
        ),

        ' + '

    ) WITHIN GROUP
    (
        ORDER BY
            revision_effective_date,
            cons_level
    ) AS corrected_calculation,


    CAST
    (
        SUM(component_amount)

        AS DECIMAL(18,2)

    ) AS corrected_amount


INTO #CorrectedConsumption


FROM #ConsumptionComponents


GROUP BY
    logical_bill_line_id;



/*==========================================================================
    10. AGGREGATE CORRECTED FLAT CHARGES
==========================================================================*/

SELECT
    logical_bill_line_id,


    STRING_AGG
    (
        CAST
        (
            'Rev '

            + CAST
              (
                  applicable_revision
                  AS VARCHAR(20)
              )

            + ' ['

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_begin,
                  101
              )

            + '-'

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_end,
                  101
              )

            + '] '

            + CAST
              (
                  revision_days
                  AS VARCHAR(10)
              )

            + '/'

            + CAST
              (
                  total_billing_days
                  AS VARCHAR(10)
              )

            + ' days; $'

            + CAST
              (
                  CAST
                  (
                      revision_minimum
                      AS DECIMAL(18,2)
                  )

                  AS VARCHAR(30)
              )

            + ' x '

            + CAST
              (
                  CAST
                  (
                      no_of_units
                      AS DECIMAL(18,3)
                  )

                  AS VARCHAR(30)
              )

            + ' units x '

            + CAST
              (
                  CAST
                  (
                      revision_ratio * 100
                      AS DECIMAL(10,2)
                  )

                  AS VARCHAR(20)
              )

            + '% = $'

            + CAST
              (
                  CAST
                  (
                      component_amount
                      AS DECIMAL(18,2)
                  )

                  AS VARCHAR(30)
              )

            AS VARCHAR(MAX)
        ),

        ' + '

    ) WITHIN GROUP
    (
        ORDER BY
            revision_effective_date
    ) AS corrected_calculation,


    CAST
    (
        SUM(component_amount)

        AS DECIMAL(18,2)

    ) AS corrected_amount


INTO #CorrectedFlat


FROM #FlatComponents


GROUP BY
    logical_bill_line_id;



/*==========================================================================
    11. CORRECTED REVISION / DAY SUMMARY
==========================================================================*/

SELECT
    logical_bill_line_id,


    STRING_AGG
    (
        CAST
        (
            'Rev '

            + CAST
              (
                  applicable_revision
                  AS VARCHAR(20)
              )

            + ' ['

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_begin,
                  101
              )

            + '-'

            + CONVERT
              (
                  VARCHAR(10),
                  revision_period_end,
                  101
              )

            + '] '

            + CAST
              (
                  revision_days
                  AS VARCHAR(10)
              )

            + '/'

            + CAST
              (
                  total_billing_days
                  AS VARCHAR(10)
              )

            + ' days ('

            + CAST
              (
                  CAST
                  (
                      revision_ratio * 100
                      AS DECIMAL(10,2)
                  )

                  AS VARCHAR(20)
              )

            + '%)'

            AS VARCHAR(MAX)
        ),

        ' | '

    ) WITHIN GROUP
    (
        ORDER BY
            revision_effective_date
    ) AS revision_breakdown


INTO #RevisionSummary


FROM #RevisionSegmentCalc


GROUP BY
    logical_bill_line_id;



/*==========================================================================
    12. BUILD FINAL LOGICAL BILL LINES

    ONE ROW PER:

        Account
        Transaction
        Rate Code
        Detail Code
        Rate Period
        Billing Period
==========================================================================*/

SELECT
    bl.logical_bill_line_id,

    bl.first_source_detail_id,

    bl.source_detail_rows,


    bl.transaction_id,

    bl.cust_no,

    bl.cust_sequence,


    bl.billing_cycle,

    bl.lot_no,

    bl.area,

    bl.st_category,


    bl.tran_date,

    bl.period_begin_date,

    bl.period_end_date,


    bl.service_description,

    bl.service_code,

    bl.detail_code,

    bl.use_period,


    bl.billed_revisions,

    bl.original_billed_breakdown,


    bl.no_of_units,

    bl.consumption,

    bl.original_billable_cons,

    bl.water_billable_cons,

    bl.winter_average,

    bl.winter_average_effective_date,

    bl.consumption_basis,

    bl.billable_cons,

    bl.original_percent_of_period,


    /*----------------------------------------------------------------------
        TRUE ORIGINAL LOGICAL LINE TOTAL
    ----------------------------------------------------------------------*/

    bl.springbrook_amount,


    rs.revision_breakdown,


    /*----------------------------------------------------------------------
        CORRECTED CALCULATION
    ----------------------------------------------------------------------*/

    CASE

        WHEN UPPER(bl.detail_code) = 'CONSUMPTION'

            THEN cc.corrected_calculation


        WHEN UPPER(bl.detail_code) = 'FLAT'

            THEN cf.corrected_calculation

    END AS corrected_calculation,


    /*----------------------------------------------------------------------
        CORRECTED TOTAL
    ----------------------------------------------------------------------*/

    CASE

        WHEN UPPER(bl.detail_code) = 'CONSUMPTION'

            THEN cc.corrected_amount


        WHEN UPPER(bl.detail_code) = 'FLAT'

            THEN cf.corrected_amount

    END AS corrected_amount


INTO #FinalLines


FROM #BillLines AS bl


LEFT JOIN #CorrectedConsumption AS cc

    ON bl.logical_bill_line_id
       = cc.logical_bill_line_id


LEFT JOIN #CorrectedFlat AS cf

    ON bl.logical_bill_line_id
       = cf.logical_bill_line_id


LEFT JOIN #RevisionSummary AS rs

    ON bl.logical_bill_line_id
       = rs.logical_bill_line_id;



/*==========================================================================
    13. BUILD TRANSACTION SUMMARY
==========================================================================*/

SELECT
    fl.cust_no,

    fl.cust_sequence,

    fl.transaction_id,


    MAX(fl.billing_cycle)
        AS billing_cycle,


    MAX(fl.lot_no)
        AS lot_no,


    MAX(fl.area)
        AS area,


    MAX(fl.st_category)
        AS st_category,


    MAX(fl.tran_date)
        AS tran_date,


    MIN(fl.period_begin_date)
        AS period_begin_date,


    MAX(fl.period_end_date)
        AS period_end_date,


    /*----------------------------------------------------------------------
        ORIGINAL BILL TOTAL

        Now based on consolidated logical billing lines.
    ----------------------------------------------------------------------*/

    CAST
    (
        SUM(fl.springbrook_amount)

        AS DECIMAL(18,2)

    ) AS original_bill_total,


    /*----------------------------------------------------------------------
        CORRECT BILL TOTAL
    ----------------------------------------------------------------------*/

    CASE

        WHEN
            COUNT(*)
            =
            COUNT(fl.corrected_amount)

        THEN
            CAST
            (
                SUM(fl.corrected_amount)

                AS DECIMAL(18,2)
            )

        ELSE NULL

    END AS correct_bill_total,


    /*----------------------------------------------------------------------
        TOTAL ADJUSTMENT
    ----------------------------------------------------------------------*/

    CASE

        WHEN
            COUNT(*)
            =
            COUNT(fl.corrected_amount)

        THEN
            CAST
            (
                SUM(fl.corrected_amount)
                -
                SUM(fl.springbrook_amount)

                AS DECIMAL(18,2)
            )

        ELSE NULL

    END AS total_adjustment,


    SUM
    (
        CASE

            WHEN fl.corrected_amount IS NULL

                THEN 1

            ELSE 0

        END

    ) AS review_line_count


INTO #TransactionSummary


FROM #FinalLines AS fl


GROUP BY
    fl.cust_no,
    fl.cust_sequence,
    fl.transaction_id;



/*==========================================================================
    14. BUILD PIVOT SOURCE

    ONE logical line now feeds the pivot.

    Example:

        WA011_CONSUMPTION_Billed
        WA011_CONSUMPTION_Correct
        WA011_CONSUMPTION_Adjustment
==========================================================================*/

SELECT
    fl.cust_no,

    fl.cust_sequence,

    fl.transaction_id,


    fl.service_code,


    UPPER(fl.detail_code)
        AS detail_code,


    v.metric_order,


    CAST
    (
        fl.service_code
        + '_'
        + UPPER(fl.detail_code)
        + '_'
        + v.metric_name

        AS VARCHAR(128)

    ) AS pivot_column,


    CAST
    (
        v.metric_value

        AS DECIMAL(18,2)

    ) AS pivot_amount


INTO #PivotSource


FROM #FinalLines AS fl


CROSS APPLY
(
    VALUES

        (
            1,
            'Billed',

            CAST
            (
                fl.springbrook_amount

                AS DECIMAL(18,2)
            )
        ),


        (
            2,
            'Correct',

            CAST
            (
                fl.corrected_amount

                AS DECIMAL(18,2)
            )
        ),


        (
            3,
            'Adjustment',

            CAST
            (
                fl.corrected_amount
                -
                fl.springbrook_amount

                AS DECIMAL(18,2)
            )
        )

) AS v
(
    metric_order,
    metric_name,
    metric_value
);



/*==========================================================================
    RESULT SET 1
    PIVOTED BILLING ADJUSTMENT SUMMARY
==========================================================================*/

DECLARE @PivotColumns NVARCHAR(MAX);

DECLARE @SQL NVARCHAR(MAX);



SELECT
    @PivotColumns =

        STRING_AGG
        (
            CAST
            (
                QUOTENAME(x.pivot_column)

                AS NVARCHAR(MAX)
            ),

            ','
        )

        WITHIN GROUP
        (
            ORDER BY
                x.service_code,
                x.detail_code,
                x.metric_order
        )


FROM
(
    SELECT DISTINCT
        service_code,
        detail_code,
        metric_order,
        pivot_column

    FROM #PivotSource

) AS x;



IF @PivotColumns IS NOT NULL

BEGIN

    SET @SQL = N'

    ;WITH PivotInput AS
    (
        SELECT
            cust_no,
            cust_sequence,
            transaction_id,

            pivot_column,
            pivot_amount

        FROM #PivotSource
    ),


    Pivoted AS
    (
        SELECT
            cust_no,
            cust_sequence,
            transaction_id,

            ' + @PivotColumns + N'

        FROM PivotInput

        PIVOT
        (
            SUM(pivot_amount)

            FOR pivot_column IN
            (
                ' + @PivotColumns + N'
            )

        ) AS p
    )


    SELECT

        RIGHT
        (
            ''000000''
            + CAST(ts.cust_no AS VARCHAR(6)),
            6
        )

        + ''-''

        +

        RIGHT
        (
            ''000''
            + CAST(ts.cust_sequence AS VARCHAR(3)),
            3

        ) AS [Account No],


        ts.billing_cycle
            AS [Billing Cycle],


        ts.lot_no
            AS [Lot No],


        ts.area
            AS [Area],


        ts.st_category
            AS [ST_Category],


        CONVERT
        (
            VARCHAR(10),
            ts.tran_date,
            101

        ) AS [Transaction Date],


        CONVERT
        (
            VARCHAR(10),
            ts.period_begin_date,
            101

        ) AS [Period Begin],


        CONVERT
        (
            VARCHAR(10),
            ts.period_end_date,
            101

        ) AS [Period End],


        DATEDIFF
        (
            DAY,
            ts.period_begin_date,
            ts.period_end_date

        ) + 1 AS [Billing Period Days],


        ' + @PivotColumns + N',


        ts.original_bill_total
            AS [Original Bill Total],


        ts.correct_bill_total
            AS [Correct Bill Total],


        ts.total_adjustment
            AS [Total Adjustment],


        CASE

            WHEN ts.total_adjustment IS NULL

                THEN ''REVIEW''


            WHEN ts.total_adjustment > 0

                THEN ''CHARGE''


            WHEN ts.total_adjustment < 0

                THEN ''CREDIT''


            ELSE ''NO ADJUSTMENT''

        END AS [Total Adjustment Type],


        CASE

            WHEN ts.total_adjustment IS NULL

                THEN NULL


            ELSE
                CAST
                (
                    ABS(ts.total_adjustment)

                    AS DECIMAL(18,2)
                )

        END AS [Total Amount to Adjust],


        CASE

            WHEN ts.review_line_count > 0

                THEN ''REVIEW''


            WHEN ts.total_adjustment > 0

                THEN ''READY - CHARGE''


            WHEN ts.total_adjustment < 0

                THEN ''READY - CREDIT''


            ELSE ''NO ADJUSTMENT''

        END AS [Adjustment Status]


    FROM #TransactionSummary AS ts


    LEFT JOIN Pivoted AS p

        ON  ts.cust_no
            = p.cust_no

        AND ts.cust_sequence
            = p.cust_sequence

        AND ts.transaction_id
            = p.transaction_id


    ORDER BY
        ts.cust_no,
        ts.cust_sequence,
        ts.tran_date;

    ';


    EXEC sys.sp_executesql @SQL;

END


ELSE

BEGIN

    SELECT
        CAST(NULL AS VARCHAR(10))
            AS [Account No],

        CAST(NULL AS INT)
            AS [Billing Cycle],

        CAST(NULL AS DECIMAL(18,2))
            AS [Total Adjustment]

    WHERE
        1 = 0;

END;



/*==========================================================================
    RESULT SET 2
    LOGICAL BILLING LINE ADJUSTMENT DETAIL

    ONE ROW PER LOGICAL BILLING LINE.

    EXAMPLE FOR ACCOUNT 000271-000:

        SW01 / Flat
        WA011 / Consumption
        WA011 / Flat

    NOT:

        WA011 Consumption Rev 22
        WA011 Consumption Rev 23

    Those original revision rows have already been combined.
==========================================================================*/

SELECT

    /*======================================================================
        ACCOUNT
    ======================================================================*/

    RIGHT
    (
        '000000'
        + CAST(cust_no AS VARCHAR(6)),
        6
    )

    + '-'

    +

    RIGHT
    (
        '000'
        + CAST(cust_sequence AS VARCHAR(3)),
        3

    ) AS [Account No],


    billing_cycle
        AS [Billing Cycle],


    lot_no
        AS [Lot No],


    area
        AS [Area],


    st_category
        AS [ST_Category],


    /*======================================================================
        BILLING PERIOD
    ======================================================================*/

    CONVERT
    (
        VARCHAR(10),
        tran_date,
        101

    ) AS [Transaction Date],


    CONVERT
    (
        VARCHAR(10),
        period_begin_date,
        101

    ) AS [Period Begin],


    CONVERT
    (
        VARCHAR(10),
        period_end_date,
        101

    ) AS [Period End],


    DATEDIFF
    (
        DAY,
        period_begin_date,
        period_end_date

    ) + 1
        AS [Total Billing Days],


    /*======================================================================
        SERVICE
    ======================================================================*/

    service_description
        AS [Service],


    service_code
        AS [Rate Code],


    detail_code
        AS [Detail Code],


    use_period
        AS [Rate Period],


    /*======================================================================
        ORIGINAL SPRINGBROOK SOURCE ROWS

        Useful for validating that duplicate revision rows were collapsed.
    ======================================================================*/

    source_detail_rows
        AS [Source Detail Rows],


    billed_revisions
        AS [Original Billed Revisions],


    original_billed_breakdown
        AS [Original Billed Breakdown],


    /*======================================================================
        UNITS / CONSUMPTION
    ======================================================================*/

    no_of_units
        AS [No. of Units],


    consumption
        AS [Consumption],


    original_billable_cons
        AS [Original Billable Consumption],


    water_billable_cons
        AS [Water Billable Consumption - Audit],


    winter_average
        AS [Winter Average - Audit],


    CONVERT
    (
        VARCHAR(10),
        winter_average_effective_date,
        101

    ) AS [Winter Average Effective Date],


    consumption_basis
        AS [Consumption Basis],


    billable_cons
        AS [Billable Consumption Used],


    original_percent_of_period
        AS [Original Total Percent of Period],


    /*======================================================================
        TRUE ORIGINAL LOGICAL LINE AMOUNT

        Sum of all original Springbrook revision rows.
    ======================================================================*/

    CAST
    (
        springbrook_amount

        AS DECIMAL(18,2)

    ) AS [Original Billed Amount],


    /*======================================================================
        CORRECT RATE REVISION BREAKDOWN
    ======================================================================*/

    revision_breakdown
        AS [Correct Revision / Day Breakdown],


    /*======================================================================
        CORRECTED CALCULATION
    ======================================================================*/

    corrected_calculation
        AS [New Calculation],


    /*======================================================================
        CORRECT AMOUNT
    ======================================================================*/

    CAST
    (
        corrected_amount

        AS DECIMAL(18,2)

    ) AS [Correct Amount],


    /*======================================================================
        ADJUSTMENT
    ======================================================================*/

    CASE

        WHEN corrected_amount IS NULL

            THEN NULL


        ELSE
            CAST
            (
                corrected_amount
                -
                springbrook_amount

                AS DECIMAL(18,2)
            )

    END AS [Adjustment Amount],


    /*======================================================================
        ADJUSTMENT TYPE
    ======================================================================*/

    CASE

        WHEN corrected_amount IS NULL

            THEN 'REVIEW'


        WHEN corrected_amount
             - springbrook_amount > 0

            THEN 'CHARGE'


        WHEN corrected_amount
             - springbrook_amount < 0

            THEN 'CREDIT'


        ELSE 'NO ADJUSTMENT'

    END AS [Adjustment Type],


    /*======================================================================
        ABSOLUTE AMOUNT TO ADJUST
    ======================================================================*/

    CASE

        WHEN corrected_amount IS NULL

            THEN NULL


        ELSE
            CAST
            (
                ABS
                (
                    corrected_amount
                    -
                    springbrook_amount
                )

                AS DECIMAL(18,2)
            )

    END AS [Amount to Adjust],


    /*======================================================================
        STATUS
    ======================================================================*/

    CASE

        WHEN corrected_amount IS NOT NULL

            THEN 'READY'


        ELSE 'CHECK RATE SETUP'

    END AS [Adjustment Status]


FROM #FinalLines


ORDER BY
    cust_no,

    cust_sequence,

    tran_date,

    service_code,

    CASE

        WHEN UPPER(detail_code) = 'CONSUMPTION'

            THEN 1


        WHEN UPPER(detail_code) = 'FLAT'

            THEN 2


        ELSE 3

    END;



/*==========================================================================
    OPTIONAL CLEANUP
==========================================================================*/

DROP TABLE IF EXISTS #WaterUsage;
DROP TABLE IF EXISTS #RawBillLines;
DROP TABLE IF EXISTS #BillLines;
DROP TABLE IF EXISTS #ServiceRevisionWindows;
DROP TABLE IF EXISTS #ConsumptionLevels;
DROP TABLE IF EXISTS #RevisionSegmentCalc;
DROP TABLE IF EXISTS #ConsumptionComponents;
DROP TABLE IF EXISTS #FlatComponents;
DROP TABLE IF EXISTS #CorrectedConsumption;
DROP TABLE IF EXISTS #CorrectedFlat;
DROP TABLE IF EXISTS #RevisionSummary;
DROP TABLE IF EXISTS #FinalLines;
DROP TABLE IF EXISTS #TransactionSummary;
DROP TABLE IF EXISTS #PivotSource;