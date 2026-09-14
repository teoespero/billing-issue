# Springbrook Latest Revisions Effective 07/01/2025

## Purpose

This SQL script finds the **highest revision number for each Springbrook service** where the service detail effective date is still **07/01/2025**.

It returns:

- Service Code
- Description
- Revision Number
- Effective Date
- Service Minimum
- Consumption Level
- Minimum Consumption
- Rate

## Author

**Teo Espero**

## Created

**09/04/2026**

## SQL Code

```sql
/*==========================================================================
    SCRIPT NAME:
        Springbrook_Latest_Revisions_Effective_07012025.sql

    PURPOSE:
        Finds the latest/highest revision number for each Springbrook
        service where the effective date is still 07/01/2025.

    AUTHOR:
        Teo Espero

    CREATED:
        09/04/2026

    TABLES USED:
        dbo.ub_service
        dbo.ub_service_detail
        dbo.ub_service_cons_lvl

    NOTES:
        - MAX(revision_no) is used to identify the highest revision number
          for each service.
        - Only service-detail records effective on 07/01/2025 are considered.
        - The date filter uses a range so it also works correctly if
          effective_date contains a time component.
==========================================================================*/

WITH LatestRevision AS
(
    SELECT
        -- Service identifier used to group revisions by service.
        sd.ub_service_id,

        -- Highest revision number for the service where the effective
        -- date is still 07/01/2025.
        MAX(sd.revision_no) AS latest_revision_no

    FROM Springbrook0.dbo.ub_service_detail AS sd

    WHERE
        sd.effective_date >= '20250701'
        AND sd.effective_date <  '20250702'

    GROUP BY
        sd.ub_service_id
)

SELECT
    -- Service information.
    s.service_code,
    s.description,

    -- Highest revision found for the service.
    sd.revision_no,

    -- Display effective date as MM/DD/YYYY.
    CONVERT(varchar(10), sd.effective_date, 101) AS effective_date,

    -- Service-level minimum.
    sd.minimum AS service_minimum,

    -- Consumption tier information.
    cl.cons_level,
    cl.minimum_cons,
    cl.rate

FROM Springbrook0.dbo.ub_service AS s

-- Connect the service to its service-detail revisions.
INNER JOIN Springbrook0.dbo.ub_service_detail AS sd
    ON s.ub_service_id = sd.ub_service_id

-- Keep only the highest revision number found for each service.
INNER JOIN LatestRevision AS lr
    ON sd.ub_service_id = lr.ub_service_id
    AND sd.revision_no = lr.latest_revision_no

-- Retrieve the consumption levels and rates for that service revision.
INNER JOIN Springbrook0.dbo.ub_service_cons_lvl AS cl
    ON sd.ub_service_detail_id = cl.ub_service_detail_id

-- Confirm that the returned revision is effective on 07/01/2025.
WHERE
    sd.effective_date >= '20250701'
    AND sd.effective_date <  '20250702'

ORDER BY
    s.service_code,
    sd.revision_no,
    cl.cons_level;
```

## How It Works

The `LatestRevision` common table expression first looks only at service-detail records with an effective date of **07/01/2025**. It groups those records by `ub_service_id` and uses `MAX(revision_no)` to identify the highest revision number for each service.

The main query then joins that result back to the Springbrook service and consumption-level tables so the corresponding service information, minimums, consumption tiers, and rates can be reviewed.
