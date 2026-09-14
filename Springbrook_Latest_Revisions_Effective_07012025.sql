/*==========================================================================
    SCRIPT NAME:
        Springbrook_Latest_Revisions_Effective_07012025.sql

    PURPOSE:
        Finds the latest/highest revision for each Springbrook service
        where the effective date is still 07/01/2025.

    AUTHOR:
        Teo Espero

    CREATED:
        09/04/2026
==========================================================================*/

WITH LatestRevision AS
(
    SELECT
        sd.ub_service_id,
        MAX(sd.revision_no) AS latest_revision_no

    FROM Springbrook0.dbo.ub_service_detail AS sd

    WHERE
        sd.effective_date >= '20250701'
        AND sd.effective_date <  '20250702'

    GROUP BY
        sd.ub_service_id
)

SELECT
    -- Service information
    s.service_code,
    s.description,

    -- Latest revision still using 07/01/2025
    sd.revision_no,

    -- Display effective date as MM/DD/YYYY
    CONVERT(varchar(10), sd.effective_date, 101) AS effective_date,

    -- Service-level minimum
    sd.minimum AS service_minimum,

    -- Consumption tier information
    cl.cons_level,
    cl.minimum_cons,
    cl.rate

FROM Springbrook0.dbo.ub_service AS s

INNER JOIN Springbrook0.dbo.ub_service_detail AS sd
    ON s.ub_service_id = sd.ub_service_id

INNER JOIN LatestRevision AS lr
    ON sd.ub_service_id = lr.ub_service_id
    AND sd.revision_no = lr.latest_revision_no

INNER JOIN Springbrook0.dbo.ub_service_cons_lvl AS cl
    ON sd.ub_service_detail_id = cl.ub_service_detail_id

WHERE
    sd.effective_date >= '20250701'
    AND sd.effective_date <  '20250702'

ORDER BY
    s.service_code,
    sd.revision_no,
    cl.cons_level;