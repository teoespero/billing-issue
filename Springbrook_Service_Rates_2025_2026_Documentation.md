# Springbrook Service Rates 2025–2026

## Overview

This SQL script is used to review and compare Springbrook Utility Billing service rates and rate revisions that became effective during **calendar years 2025 and 2026**.

It is intended to help verify that the correct service rate revisions are present in Springbrook and to make it easier to compare older and newer rate structures during billing review, rate validation, testing, and troubleshooting.

## What the Script Is For

The script pulls service rate information from the Springbrook database so staff can see which rates were associated with each service revision.

It is especially useful for:

- Reviewing service rates before running billing
- Comparing 2025 and 2026 rate revisions
- Verifying that the correct revision numbers are assigned
- Confirming effective dates
- Reviewing minimum charges and consumption tiers
- Comparing minimum consumption values and tier rates
- Supporting billing validation in TEST or LIVE
- Troubleshooting incorrect billing caused by rate or revision problems
- Exporting rate information to Excel for additional review

## Information Returned

The query returns the following fields:

| Field | Description |
|---|---|
| `service_code` | Springbrook service code |
| `description` | Description of the service |
| `revision_no` | Revision number assigned to the service rate configuration |
| `effective_date` | Date the revision became effective, displayed as MM/DD/YYYY |
| `service_minimum` | Minimum value or charge stored at the service revision level |
| `cons_level` | Consumption tier or level |
| `minimum_cons` | Minimum consumption associated with the tier |
| `rate` | Rate charged for that consumption level |

## Database and Tables

**Database:** `Springbrook0`

The script uses three Springbrook Utility Billing tables:

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

### `ub_service`

Contains the main service information, including the service code and description.

### `ub_service_detail`

Contains the service revision information, including:

- Revision number
- Effective date
- Minimum service value
- Other rate configuration settings

A single service may have several service-detail records because Springbrook keeps separate revisions of a service configuration.

### `ub_service_cons_lvl`

Contains the individual consumption levels or rate tiers associated with each service-detail revision.

This table provides the tier number, minimum consumption, and rate.

## How the Script Works

The query starts with the main service table and joins it to the service-detail table using `ub_service_id`.

It then joins the service-detail record to the consumption-level table using `ub_service_detail_id`.

This allows the query to connect:

**Service → Revision → Consumption Tier → Rate**

The query filters the results to revisions with effective dates between:

- **01/01/2025**
- **12/31/2026**

The filter is written as:

```sql
sd.effective_date >= '20250101'
AND sd.effective_date < '20270101'
```

Using an exclusive upper boundary ensures that all records through December 31, 2026 are included, even if the field contains a time value.

## Date Formatting

The effective date is displayed using:

```sql
CONVERT(varchar(10), sd.effective_date, 101)
```

This displays the date in:

```text
MM/DD/YYYY
```

Example:

```text
07/01/2025
07/01/2026
```

## Result Sorting

The results are sorted by:

1. Service Code
2. Effective Date
3. Revision Number
4. Consumption Level

This makes it easier to compare multiple revisions and rate tiers for the same service.

## Practical Use

For rate validation, the output can be exported to Excel and reviewed side-by-side.

A typical review would compare:

```text
Service Code
    ↓
2025 Revision
    ↓
2025 Consumption Levels and Rates
    ↓
2026 Revision
    ↓
2026 Consumption Levels and Rates
```

This can help identify issues such as:

- A service still using an older effective date
- A revision number that does not match expectations
- Missing consumption tiers
- Incorrect minimum consumption
- Incorrect rate values
- A newer rate structure attached to the wrong revision
- An older rate structure recreated under a newer revision

## Important Notes

- `sd.minimum` is the minimum value stored at the **service revision level**.
- `cl.minimum_cons` is the minimum consumption stored at the **consumption tier level**.
- These are different fields and both can be useful when reviewing rate configurations.
- The script is read-only and does not modify Springbrook data.
- The query should be run against the appropriate Springbrook database environment depending on whether the review is being performed in TEST or LIVE.

## Authorship

**Author:** Teo Espero  
**Original Creation Date:** 09/04/2026

## Revision History

| Version | Date | Author | Description |
|---|---|---|---|
| 1.0 | 09/04/2026 | Teo Espero | Initial query created to retrieve service rates and revisions for 2025 and 2026 |
| 1.1 | 09/04/2026 | Teo Espero | Added MM/DD/YYYY date formatting |
| 1.2 | 09/04/2026 | Teo Espero | Added documentation, aliases, revision history, and inline comments |
| 1.3 | 09/09/2026 | Teo Espero | Added standalone Markdown documentation describing the script's purpose and use |

## Summary

The purpose of this script is to provide a clear view of Springbrook service rate revisions for 2025 and 2026 so the rates can be reviewed, compared, and validated before or during billing operations.

It is primarily a **rate verification and troubleshooting tool** and does not make any changes to the Springbrook database.
