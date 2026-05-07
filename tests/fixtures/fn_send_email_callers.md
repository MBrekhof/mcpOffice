# FN_SEND_EMAIL Callers — Contact Registration Audit

> Verified against export `be5aef3` (2026-05-01). Scratchpad routines (`_*.vb`) excluded.

`FN_SEND_EMAIL` has an **optional** registration mechanism (lines 359–367):

```
IF (emailSend) THEN 
    IF (NotEmpty(contactUpdateNumber)) THEN 
        UPDATE CONTACT SET MEMO = 'Email send -<date>-' WHERE CONTACT_NUMBER = {contactUpdateNumber}
    ENDIF 
    contactUpdateNumber = NullValue() 
ENDIF 
```

To register an email send in the CONTACT table, the caller must:
1. Create the CONTACT row first via `GOSUB FN_NEW_CONTACT` (sets `contactNumber`)
2. Assign `contactUpdateNumber = contactNumber`
3. Pass `contactUpdateNumber` to `FN_SEND_EMAIL` (either directly via context, or as a `variableNames`/`variableValues` entry when using `BackgroundSubroutine`)

This audit categorizes all 36 production callers by whether they have this wiring.

---

## Category A — Wired correctly (registration works) — 2 files

| Caller | Email type | Mechanism | Line |
|--------|-----------|-----------|------|
| `SA_LABWARE_TO_SAMPLEAPP` | SampleApp batch failure notice | BackgroundSubroutine | 852 |
| `FN_CNTRL_CHART_INVESTIGATION` | Control chart out-of-trend alert | BackgroundSubroutine | 247 |

Both call `FN_NEW_CONTACT`, assign `contactUpdateNumber = contactNumber`, and pass it through `variableNames`/`variableValues`.

---

## Category B — Has FN_NEW_CONTACT but DOES NOT propagate contactUpdateNumber — 1 file

These callers create a CONTACT row but never register the email send against it. Each one is a candidate for a one-line fix (`contactUpdateNumber = contactNumber` after the `GOSUB FN_NEW_CONTACT`).

| Caller | Email type | FN_NEW_CONTACT | Send call |
|--------|-----------|----------------|-----------|
| `FN_OOS_MAILING_CHECK` | OOS-specification breach alert | line 729 | line 731 |

**Fix:** between lines 729 and 731 of `FN_OOS_MAILING_CHECK.vb`, add:

```
contactUpdateNumber = contactNumber
```

The CONTACT row is already created (line 729). One line wires up the registration.

---

## Category C — Calls FN_SEND_EMAIL without FN_NEW_CONTACT — 33 files

These callers do not create a CONTACT row at all. To register, both `FN_NEW_CONTACT` and `contactUpdateNumber` would need to be added. Whether to do so is a per-routine business decision (some emails are clearly transactional and don't warrant a CONTACT record; others probably do).

### Likely candidates for adding registration

Routines whose emails relate to a customer, project, or sample workflow — adding a CONTACT entry would create an audit trail:

| Caller | Email type | Send line |
|--------|-----------|-----------|
| `FN_BM_NEXT_STATUS` | Batch status transition notification | GOSUB 190 |
| `FN_CHANGE_TEMPLATE_NOTIFY` | Template change notification | 227 |
| `FN_CHARGE_REQUEST_UPDATE_STATUS` | Charge request status update | 171 |
| `FN_CONTACT_SEND_MESSAGE` | Contact-driven outbound message | 237 — **manages own CONTACT bookkeeping via SQL (lines 55, 261)** — different pattern, already registers |
| `FN_DUPLO_CHECK` | Duplicate-result alert | 114 |
| `FN_EMAIL_GEN_CHECK` | Generic email rule check (08:00 daily) | GOSUB 473 |
| `FN_OOS_SEND_EMAIL` | OOS notification dispatcher | 161 |
| `FN_ORDER_SEND_EMAIL` | Order-related notification | 101 |
| `FN_OVERDUE_ANALYSIS` | Overdue analysis report | **commented out 99** |
| `FN_SENDSCHEDULE_EMAIL` | Sample-schedule email (2 calls) | GOSUB 114, 119 |
| `FN_SEND_INVOICE_MB` | Invoice email | 141 |
| `FN_SHIP_SAMPLES_SEND` | Shipping notification | 113 |
| `IR_RESULT_TREND_BACKGROUND` | Result trend chart email (deferred per migration plan) | 260 |
| `INV_FN_CHECK_REORDER_LEVEL_AND_SEND_EMAIL` | Inventory reorder alert | 170 |
| `ME_ALERT_BY_MAIL` | Message Engine alert wrapper | 73 |
| `VWF_CUST_WEB_MAINTENANCE_LAUNCHER` | Customer web maintenance (2 calls) | GOSUB 268, 294 |
| `VWF_TNATRENDING_SENDMAIL` | TAT trending email | 57 |

### Likely intentional non-registration

Routines whose emails are operational/internal and don't need CONTACT records:

| Caller | Email type | Send line |
|--------|-----------|-----------|
| `ET_EXP_CADRI_VITENS_RES` | CADRI export error notification | 180 |
| `FN_BIO_PRINT_LBL_SCHED_SAMP` | Biology label print error | 104 |
| `FN_CHECK_SQL_INPUT` | SQL injection lint alert | GOSUB 78 |
| `FN_CNTRLCHRT_EVAL_CHRT` (×3) | Control chart eval failures | 210, 226, 298 |
| `FN_CREATE_INSTRUMENT_EXPORT` | Instrument export error | 291 |
| `FN_CREATE_SCE_BGS` | SCE BGS export error | 86 |
| `FN_CREATE_SKORF_BGS` | SKORF BGS export error | 92 |
| `FN_CREATE_TACC_EXPORT` | TACC export error | 208 |
| `FN_EXP_METALS_BATCH` | Metals batch export error | 69 |
| `FN_EXP_OPD_CADRI_VITENS` | CADRI Vitens export error | 131 |
| `FN_EXP_OPD_CADRI_VITENS_RK` | CADRI Vitens RK export error | 118 |
| `ME_PROC_AQUO_CSV` (×2) | Message Engine: AQUO CSV processor | 89, 115 |
| `ME_PROC_CAMSIZER_PDF` | Message Engine: Camsizer PDF | 157 |
| `ME_PROC_PGIM_STATUS` | Message Engine: PGIM status | 171 |
| `SCHED_IM_EXP_DAILY` | Daily IM export | 136 |
| `SCHED_IM_EXP_MONTHLY` | Monthly IM export | 147 |
| `UTIL_LIMS_LOG_ACTION_SEND_EMAIL` | LIMS_LOG action email | 178 |
| `FN_ADD_ORGANISM` (SYSTEM) | Organism added notification | 94 |
| `FN_INTERFERENCE_COMPLETE` (SYSTEM) | Interference test complete | 329 |
| `FN_XML_MARTINI` (SYSTEM) | Martini XML export | 153 |
| `ME_PROC_WAARDENBURG` (SYSTEM) | Message Engine: Waardenburg | 211 |

---

## Summary

| Category | Count | Action |
|----------|-------|--------|
| A — wired correctly | 2 | none |
| B — has CONTACT row, doesn't propagate | 1 | one-line fix in `FN_OOS_MAILING_CHECK` |
| C(a) — likely candidates for adding registration | 16 routines (18 send sites) | per-routine business decision |
| C(b) — likely intentional non-registration | 17 routines (21 send sites) | leave as-is |
| C(special) — `FN_CONTACT_SEND_MESSAGE` | 1 | already registers via its own SQL bookkeeping |

The cheapest, highest-value action is the one-line fix in `FN_OOS_MAILING_CHECK` (Category B). Beyond that, decide per routine whether the email warrants a CONTACT audit trail.

## How to verify "is this email registered?"

For any caller in Category C, run this SQL in Labware8 to see whether equivalent CONTACT records exist for past sends:

```sql
SELECT TOP 50 CONTACT_NUMBER, CREATED_ON, MEMO 
  FROM CONTACT 
 WHERE MEMO LIKE 'Email send -%' 
 ORDER BY CREATED_ON DESC
```

If there are very few such rows relative to the actual email volume, the registration mechanism is being underused.
