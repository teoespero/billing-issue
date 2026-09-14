# Springbrook Service Rates 2025–2026

## Script Information

**Script Name:** `Springbrook_Service_Rates_2025_2026.sql`  
**Author:** Teo Espero  
**Created:** 09/04/2026  
**Database:** `Springbrook0`

## Purpose

This SQL script retrieves Springbrook Utility Billing service rates and rate revisions with effective dates in calendar years 2025 and 2026.

The report includes:

- Service Code
- Service Description
- Revision Number
- Effective Date
- Service Minimum
- Consumption Level
- Minimum Consumption
- Rate

## Tables Used

- `dbo.ub_service`
- `dbo.ub_service_detail`
- `dbo.ub_service_cons_lvl`

## Table Relationships

```text
ub_service
    |
    | ub_service_id
    v
ub_service_detail
    |
    | ub_service_detail_id
    v
ub_service_cons_lvl
```

## Revision History

| Version | Date | Author | Description |
|---|---|---|---|
| 1.0 | 09/04/2026 | Teo Espero | Initial query created to retrieve service rates and revisions for 2025 and 2026. |
| 1.1 | 09/04/2026 | Teo Espero | Added MM/DD/YYYY formatting for effective dates. |
| 1.2 | 09/04/2026 | Teo Espero | Added documentation, revision history, aliases, and inline comments. |

## Notes

1. The query only returns service-detail revisions with effective dates from 01/01/2025 through 12/31/2026.
2. `sd.minimum` represents the minimum value stored at the service revision level.
3. `cl.minimum_cons` represents the minimum consumption associated with each consumption level or tier.
4. Effective Date is converted to `MM/DD/YYYY` for easier reporting and spreadsheet review.
5. The `WHERE` clause uses an exclusive upper date boundary instead of `YEAR(effective_date)`, which can help SQL Server use an index on `effective_date` more efficiently if one exists.
6. SQL Server-safe `YYYYMMDD` date literals are used in the filter to avoid regional date-format issues.

## Final SQL Code

```sql
/*==========================================================================

    SCRIPT NAME:
        Springbrook_Service_Rates_2025_2026.sql

    PURPOSE:
        Retrieves Springbrook Utility Billing service rates and rate
        revisions with effective dates in calendar years 2025 and 2026.

        The report includes:
            - Service Code
            - Service Description
            - Revision Number
            - Effective Date
            - Service Minimum
            - Consumption Level
            - Minimum Consumption
            - Rate

    DATABASE:
        Springbrook0

    TABLES USED:
        dbo.ub_service
        dbo.ub_service_detail
        dbo.ub_service_cons_lvl

    AUTHOR:
        Teo Espero

    CREATED:
        09/04/2026

    ------------------------------------------------------------------------
    REVISION HISTORY
    ------------------------------------------------------------------------
    Version     Date          Author          Description
    ------------------------------------------------------------------------
    1.0         09/04/2026    Teo Espero      Initial query created to
                                               retrieve service rates and
                                               revisions for 2025 and 2026.

    1.1         09/04/2026    Teo Espero      Added MM/DD/YYYY formatting
                                               for effective dates.

    1.2         09/04/2026    Teo Espero      Added documentation, revision
                                               history, aliases, and inline
                                               comments.

    ------------------------------------------------------------------------
    TABLE RELATIONSHIPS
    ------------------------------------------------------------------------

        ub_service
            |
            | ub_service_id
            v
        ub_service_detail
            |
            | ub_service_detail_id
            v
        ub_service_cons_lvl

    ------------------------------------------------------------------------
    NOTES
    ------------------------------------------------------------------------
    1. The query only returns service-detail revisions with effective dates
       from 01/01/2025 through 12/31/2026.

    2. sd.minimum represents the minimum value stored at the service
       revision level.

    3. cl.minimum_cons represents the minimum consumption associated with
       each consumption level/tier.

    4. Effective Date is converted to MM/DD/YYYY for easier reporting and
       spreadsheet review.

    5. The WHERE clause uses an exclusive upper date boundary rather than
       YEAR(effective_date). This allows SQL Server to use an index on
       effective_date more efficiently if one exists.

==========================================================================*/


SELECT

    /*----------------------------------------------------------------------
      SERVICE INFORMATION
      Comes from dbo.ub_service.
    ----------------------------------------------------------------------*/

    s.service_code AS service_code,

    -- Description/name of the Springbrook service.
    s.description AS description,


    /*----------------------------------------------------------------------
      RATE REVISION INFORMATION
      Comes from dbo.ub_service_detail.

      Each service may have multiple revisions with different effective
      dates and rate structures.
    ----------------------------------------------------------------------*/

    -- Springbrook revision number for this service configuration.
    sd.revision_no AS revision_no,

    -- Display effective date as MM/DD/YYYY.
    CONVERT(varchar(10), sd.effective_date, 101) AS effective_date,

    -- Minimum amount/value defined at the service revision level.
    sd.minimum AS service_minimum,


    /*----------------------------------------------------------------------
      CONSUMPTION LEVEL / TIER INFORMATION
      Comes from dbo.ub_service_cons_lvl.

      A single service revision may contain multiple consumption levels.
    ----------------------------------------------------------------------*/

    -- Consumption tier/level number.
    cl.cons_level AS cons_level,

    -- Minimum consumption associated with this consumption level.
    cl.minimum_cons AS minimum_cons,

    -- Rate charged for this consumption level.
    cl.rate AS rate


FROM Springbrook0.dbo.ub_service AS s


/*--------------------------------------------------------------------------
  JOIN SERVICE TO SERVICE DETAIL

  ub_service_id identifies the parent service record.

  One service can have multiple service-detail records representing
  different revisions/effective dates.
--------------------------------------------------------------------------*/

INNER JOIN Springbrook0.dbo.ub_service_detail AS sd
    ON s.ub_service_id = sd.ub_service_id


/*--------------------------------------------------------------------------
  JOIN SERVICE DETAIL TO CONSUMPTION LEVELS

  ub_service_detail_id identifies the specific service revision to which
  the consumption level/rate belongs.
--------------------------------------------------------------------------*/

INNER JOIN Springbrook0.dbo.ub_service_cons_lvl AS cl
    ON sd.ub_service_detail_id = cl.ub_service_detail_id


/*--------------------------------------------------------------------------
  FILTER

  Return only revisions effective during calendar years 2025 and 2026.

  Using:
      >= 01/01/2025
      <  01/01/2027

  ensures all dates through 12/31/2026 are included, including values where
  effective_date may contain a time component.
--------------------------------------------------------------------------*/

WHERE
    sd.effective_date >= '20250101'
    AND sd.effective_date < '20270101'


/*--------------------------------------------------------------------------
  SORT ORDER

  Results are grouped logically by:
      1. Service Code
      2. Effective Date
      3. Revision Number
      4. Consumption Level

  This makes comparing 2025 and 2026 rates easier.
--------------------------------------------------------------------------*/

ORDER BY
    s.service_code,
    sd.effective_date,
    sd.revision_no,
    cl.cons_level;
```
