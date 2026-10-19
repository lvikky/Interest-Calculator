# Python Data Backfill Automation

**Code-verified explanation and operational walkthrough**  
**Source reviewed:** `Pasted text(7).txt` (420-line Python script)  
**Prepared:** 19 October 2026  
**Audience:** Technical and non-technical readers

> **Evidence standard:** Everything described as *verified* is supported by the supplied Python script. Examples are explicitly hypothetical. The actual INI file, live database schema, real records, logs, and successful execution evidence were not supplied. This document does not claim to have inspected them.

## 1. Executive overview

The script performs a **date-based backfill within one SQL Server table**. In simple terms, it finds existing records for a selected **source date** and can insert copies into the **same table** with a requested **target date**. It builds the SQL from the table's column metadata, performs checks before writing, and requires an explicit `--EXECUTE` option to send the `INSERT` statement.

It supports one approved table and one source/target date pair per run. It does **not** read data from a CSV, Excel file, external API, or different source table. The generated SQL reads and writes the same three-part SQL Server table name. **Evidence:** `build_insert_sql()` (lines 209–227); `run_backfill()` (lines 277–372).

### A simple analogy

Imagine a cupboard containing folders labelled by date. The script can look at folders marked **Date A**, check that there are no folders for **Date B**, and prepare copies of the Date A folders relabelled Date B. It first tells you what it plans to do. Only when you explicitly enable execution does it try to add the new folders.

**What this analogy does not imply:** The Python script does not know your business reasons for copying the records, and it does not validate that each copied record is correct at the business level.

## 2. Verified data sources, destination, and boundaries

| Element | What the code establishes | Source |
|---|---|---|
| Existing business rows | Read from the selected table using `WHERE <date_column> = ?`. | Lines 209–227 |
| New business rows | Inserted back into that **same** selected table using `INSERT INTO ... SELECT`. | Lines 209–227 |
| Table identity | Built as `LIQUIDITY.LST_SF_EXTRACT.<TABLE>`. | Lines 27–29, 112–113 |
| SQL Server connection | Created with `pyodbc`; server, initial database, and connection options come from an INI section. | Lines 47–100, 289–290 |
| Table existence metadata | Obtained from `INFORMATION_SCHEMA.TABLES`. | Lines 129–142 |
| Column metadata | Obtained from `sys.columns`, `sys.tables`, `sys.schemas`, and `sys.types`. | Lines 145–178 |
| Runtime input | Arguments passed on the command line. | Lines 394–402 |
| Other data sources | No file-based import, API download, or additional business-row source table appears in the script. | Full supplied source |

**Important distinction:** The initial `DATABASE` value in the ODBC connection is read from the INI file. The row-count and `INSERT` queries use the fixed database name `LIQUIDITY`, but the catalog queries do **not** explicitly name that database. As the INI contents were not supplied, this document cannot confirm whether the configured initial database is `LIQUIDITY`. **Evidence:** lines 57–60, 112–113, 129–160, 209–235.

## 3. Supported tables and date rules

The script accepts exactly these 12 table names through a hard-coded allowlist:

| # | Approved table | Source filter/date overwritten |
|---|---|---|
| 1 | `BANKING_COMMERCIAL_PAPER` | `AS_OF_DATE` |
| 2 | `BANKING_DEPOSITS` | `AS_OF_DATE` |
| 3 | `BANKING_LOANS` | `AS_OF_DATE` |
| 4 | `BANKING_SFTS` | `AS_OF_DATE` |
| 5 | `CM_DERIVS_CASHFLOWS` | `AS_OF_DATE` |
| 6 | `CM_NOTES` | `AS_OF_DATE` |
| 7 | `CM_SFTS` | `AS_OF_DATE` |
| 8 | `CM_SLEEPER_COLLATERAL` | `AS_OF_DATE` |
| 9 | `CM_TENDER_OPTION_BONDS` | `AS_OF_DATE` |
| 10 | `CM_TOTAL_RETURN_SWAPS` | `AS_OF_DATE` |
| 11 | `DERIVS_COLLATERAL` | `AS_OF_DATE` |
| 12 | `NET_LIABILITIES` | Filter and overwrite `POSDATE`; **also** overwrite `DTPOS` |

For `NET_LIABILITIES`, the `WHERE` condition is based on `POSDATE`, **not** on `DTPOS`. Both `POSDATE` and `DTPOS` receive the target date in new rows. Other included columns are copied without date-specific substitution. **Evidence:** `APPROVED_TABLES` (lines 12–25), `get_date_column_for_table()` (lines 116–119), `build_insert_sql()` (lines 209–227).

## 4. Command-line options

| Option | Required | Code-defined meaning |
|---|---|---|
| `--ENVIRONMENT` | Yes | Exactly `DEV`, `UAT`, or `PROD` in the command-line parser. |
| `--TABLE` | Yes | Exactly one of the approved names above. |
| `--SOURCE_DATE` | Yes | Existing-record selection date, supplied in `YYYY-MM-DD` format. |
| `--TARGET_DATE` | Yes | Date assigned to the new record(s), supplied in `YYYY-MM-DD` format. |
| `--CONFIG_FILE` | No | Path to INI configuration; defaults to `dbconfig.ini`. |
| `--EXECUTE` | No | Presence enables attempted database insert after all checks. Absence means preview only. |

The script parses dates using `datetime.strptime(value, "%Y-%m-%d")`, so impossible calendar dates fail. It rejects equal source and target dates. **It does not require the target date to be later than the source date** and does not implement a working-day calendar check. **Evidence:** lines 103–109, 277–285, 394–402.

### Hypothetical invocation

The uploaded file was provided as a `.txt` attachment, so the actual Python filename is **not established**. In the examples below, `your_script.py` is a placeholder and the dates are imaginary.

**Preview-only example:**

```bash
python your_script.py --ENVIRONMENT DEV --TABLE BANKING_LOANS --SOURCE_DATE 2026-09-17 --TARGET_DATE 2026-09-18
```

**Execution-enabled example:**

```bash
python your_script.py --ENVIRONMENT DEV --TABLE BANKING_LOANS --SOURCE_DATE 2026-09-17 --TARGET_DATE 2026-09-18 --EXECUTE
```

The script changes the process's current directory to the directory containing the script before parsing and running. A relative `dbconfig.ini` path therefore resolves relative to the script directory. **Evidence:** lines 418–420.

## 5. Connecting to SQL Server

The script looks for an INI section named `SQLSERVER_DEV`, `SQLSERVER_UAT`, or `SQLSERVER_PROD`, based on the selected environment. **Evidence:** lines 47–55.

| INI key | Code behavior |
|---|---|
| `DRIVER` | Mandatory; name of the preferred ODBC driver. |
| `SERVER` | Mandatory; SQL Server address/instance value. |
| `DATABASE` | Mandatory; connection's initial database. |
| `Trusted_Connection` | Optional; defaults to `yes`. |
| `Encrypt` | Optional; defaults to `no`. |
| `TrustServerCertificate` | Optional; defaults to `yes`. |

The code checks installed drivers using `pyodbc.drivers()`. If the configured driver is not installed, it tries, in order, `ODBC Driver 18 for SQL Server`, `ODBC Driver 17 for SQL Server`, and `SQL Server`. It raises an error if no supported listed driver is present. The connection is created with `autocommit=False` and `timeout=120`. **Evidence:** lines 65–100, 289–290.

**Do not infer:** The actual server names, login identity, permission levels, certificate settings in real environments, and the actual contents of `dbconfig.ini` are unknown. The defaults listed here are **code observations, not security recommendations**.

## 6. End-to-end workflow

The following text flowchart follows the code, not an assumed business process:

```text
START
  |
  v
Parse command-line arguments and dates
  |
  v
Verify dates differ; inspect INI config and ODBC driver
  |
  v
Open SQL Server connection; inspect table and column metadata
  |
  v
Choose date column; determine insertable columns; generate SQL
  |
  v
Count SOURCE_DATE rows and TARGET_DATE rows
  |
  v
Does source count > 0 AND target count = 0?
  |                             |
  | Yes                         | No
  v                             v
Print preview + generated SQL   Error path (no INSERT)
  |
  v
Was --EXECUTE provided?
  |                       |
  | Yes                   | No
  v                       v
Run INSERT ... SELECT    Finish without INSERT
  |
  v
Count target rows again on the same connection
  |
  v
Is target count >= original source count?
  |                         |
  | Yes                     | No
  v                         v
COMMIT + success            Error; attempt ROLLBACK
```

The rollback path shown here applies to errors caught inside `run_backfill()`'s connection-handling `try` block. Failures before that block are handled by the outer `main()` exception handler instead. **Evidence:** lines 277–420.

## 7. Detailed pre-insert safety checks

These conditions are checked in the observed order before the database insert:

| Order | Validation | What happens if it fails |
|---|---|---|
| 1 | Environment is supported and both dates can be parsed; dates are different. | Stop before connecting. |
| 2 | INI path exists; environment section and required keys can be used; a supported driver exists. | Stop during setup. |
| 3 | Schema/table is found in `INFORMATION_SCHEMA.TABLES`. | Raise table-not-found error. |
| 4 | Column metadata is returned. | Raise no-metadata error. |
| 5 | Primary date column exists; additionally `DTPOS` exists for `NET_LIABILITIES`. | Raise missing-column error. |
| 6 | Primary date column is insertable; for `NET_LIABILITIES`, `DTPOS` is insertable. | Raise insertable-column error. |
| 7 | Source-date row count is greater than zero. | Stop; no source rows available to copy. |
| 8 | Target-date row count is **zero**. | Stop; refuses to add to an already populated target date. |

The checks inspect existing table metadata and row counts; **they do not establish** that every relevant SQL Server table constraint will permit the eventual insert. **Evidence:** lines 277–337.

### What counts as an insertable column?

The script reads six metadata fields: `COLUMN_NAME`, `ORDINAL_POSITION`, `DATA_TYPE`, `IS_COMPUTED`, `IS_IDENTITY`, and `IS_ROWVERSION`. It excludes computed, identity, and SQL Server `timestamp`/rowversion columns from the generated insert. It retains the other returned columns. If identity columns exist, a warning says their values will not be copied. **Evidence:** lines 145–206.

The code uses metadata ordered by `column_id`; it does not hard-code the ordinary business column names in the insert statement.

## 8. How the SQL is generated

The core operation is `INSERT INTO <same_table> (...) SELECT ... FROM <same_table> WHERE <source_date_column> = ?`.

- The `INSERT` column list includes every metadata-derived insertable column.
- For most tables, the `AS_OF_DATE` expression is replaced by `? AS AS_OF_DATE`.
- For `NET_LIABILITIES`, both `POSDATE` and `DTPOS` expressions are replaced by parameter placeholders.
- All other insertable-column expressions select the existing source-column values.
- The `WHERE` clause filters by the table's primary date column using the source date.

**Evidence:** lines 181–227.

### Structural example only: ordinary table

The next snippet is **not copied SQL output from a live table**; `[OTHER_COLUMN]` stands for any actual insertable column returned from the catalog.

```sql
INSERT INTO LIQUIDITY.LST_SF_EXTRACT.BANKING_LOANS
    (AS_OF_DATE, [OTHER_COLUMN])
SELECT
    ? AS AS_OF_DATE,
    [OTHER_COLUMN]
FROM LIQUIDITY.LST_SF_EXTRACT.BANKING_LOANS
WHERE AS_OF_DATE = ?;
```

**Bound parameter sequence:** `(TARGET_DATE, SOURCE_DATE)` for ordinary tables, assuming their only date replacement is `AS_OF_DATE`.

### Structural example only: `NET_LIABILITIES`

```sql
INSERT INTO LIQUIDITY.LST_SF_EXTRACT.NET_LIABILITIES
    (POSDATE, DTPOS, [OTHER_COLUMN])
SELECT
    ? AS POSDATE,
    ? AS DTPOS,
    [OTHER_COLUMN]
FROM LIQUIDITY.LST_SF_EXTRACT.NET_LIABILITIES
WHERE POSDATE = ?;
```

**Bound parameter sequence:** `(TARGET_DATE, TARGET_DATE, SOURCE_DATE)` in the simplified column-order illustration. In real execution, the script builds placeholder values in the order those date columns occur in the metadata-derived insertable-column list. **Evidence:** lines 209–227, 349–360.

The square-bracket column label above is illustrative only; **the generated script SQL does not automatically add brackets around its real column names**. The actual ordering and names depend on the runtime metadata, which has not been supplied.

## 9. Preview mode versus execute mode

| Operation | Without `--EXECUTE` | With `--EXECUTE` |
|---|---|---|
| Read INI and establish SQL Server connection | Yes | Yes |
| Inspect metadata and count source/target rows | Yes | Yes |
| Print details and generated SQL | Yes | Yes |
| Send the generated `INSERT` to SQL Server | No | Yes, after checks |
| Recount target rows after insert | No | Yes |
| Commit inserted records | No | Yes, if the count validation succeeds |

The preview prints environment, table, source/target dates, row counts, insertable-column count, generated SQL, and parameters. **Evidence:** lines 246–274, 338–372.

**Easy-to-miss detail:** `print_preview()` is called *before* the `if args.EXECUTE:` branch in both modes. Therefore, even in an execution-enabled run, it initially prints `MODE: PREVIEW` and `NO DATABASE CHANGES HAVE BEEN MADE.` This is true **at the moment of printing**; it does not mean an execution-enabled run will remain read-only. **Evidence:** lines 338–350.

## 10. Commit, rollback, validation, and error handling

The connection uses `autocommit=False`. In execute mode, the script sends one generated `INSERT ... SELECT` using bound date parameters. It then runs another target-date count query **on the same connection, before commit**. If the count is not below the original source count, the script commits and prints success. **Evidence:** lines 289–290, 349–372.

### The exact post-insert rule

```python
if inserted_target_count < source_count:
    raise ValueError(...)
```

That means `inserted_target_count >= source_count` **passes** this particular check. It is **not** an exact count-equality rule, a per-row comparison, a checksum, or a guarantee that every value in each copied row matches the original. **Evidence:** lines 360–369.

| Hypothetical source count | Hypothetical target count after insert | Result of this exact check |
|---:|---:|---|
| 120 | 119 | Fails because 119 is below 120. |
| 120 | 120 | Passes. |
| 120 | 121 | Also passes this count check. |

These are **mathematical examples**, not observed database results.

### Distinct error paths

| When the error happens | Observed handling |
|---|---|
| During the inner `try` (e.g., table check, source count check, `INSERT`, post-insert validation) | Log error/traceback, attempt `conn.rollback()`, print `Operational error`, return `1`, attempt cursor/connection close in `finally`. |
| Before the inner `try` (e.g., environment/date setup, config validation, creating connection/cursor) | Error can propagate to `main()`, which logs/prints `Fatal error` and returns `1`. |
| During argument parsing | `argparse` handles missing/invalid command-line options before the `main()` exception block. |
| During cleanup inside the final block | Exceptions from `cursor.close()` and `conn.close()` are suppressed. |

The success execution branch returns `0`. A successful preview-only flow does not contain an explicit return in `run_backfill()` and therefore yields `None`, which gives a zero exit status when passed to `sys.exit(main())`. **Evidence:** lines 373–420.

## 11. Hypothetical worked examples

### Example A: ordinary-table backfill

**Assumptions solely for illustration:** `BANKING_LOANS` contains 120 source rows for `2026-09-17` and zero target rows for `2026-09-18`; all catalog checks and SQL constraints are satisfied.

1. The script reads `BANKING_LOANS` metadata and identifies `AS_OF_DATE` as the relevant date column.
2. The source count is hypothetically 120; the target count is 0, so both pre-insert row-count checks pass.
3. The preview shows the proposed `INSERT ... SELECT` and parameters.
4. Without `--EXECUTE`, there is no insert.
5. With `--EXECUTE`, it attempts to insert 120 rows with the new `AS_OF_DATE`, retaining other included source values.
6. If the post-insert target count becomes 120, the count check passes and the code commits.

**Evidence for the mechanism:** lines 116–119, 209–227, 320–371. **Not evidence:** the illustrative dates, 120-row figure, or an actual successful run.

### Example B: `NET_LIABILITIES` backfill

**Assumptions solely for illustration:** `NET_LIABILITIES` has 35 source rows where `POSDATE = 2026-09-17`, zero rows at target `POSDATE = 2026-09-18`, and an insertable `DTPOS` column.

| Row element | What happens in new rows |
|---|---|
| Which source rows are selected? | `POSDATE = 2026-09-17`. |
| New `POSDATE` | Set to `2026-09-18`. |
| New `DTPOS` | Also set to `2026-09-18`. |
| Other insertable columns | Values taken from the source row. |
| Computed/identity/rowversion fields | Excluded from the explicit insert list. |

There is **no** additional source-row filter on `DTPOS` in the code, and the inserted `DTPOS` value is replaced regardless of its source-row value. **Evidence:** lines 116–119, 181–227, 349–360.

### Example C: refusals before insert

| Hypothetical condition | What the script does |
|---|---|
| 0 source rows, 0 target rows | Raises the source-zero error before printing the preview. |
| 120 source rows, 3 target rows | Raises the target-populated error; no insert attempted. |
| Source and target dates are identical | Raises an error before connecting. |
| `NET_LIABILITIES` has no insertable `DTPOS` | Raises an error before building/executing the insert. |

**Evidence:** lines 284–285, 308–337.

## 12. Function-by-function reference

| Function or entry point | Source lines | Verified responsibility |
|---|---|---|
| `setup_logger()` | 32–41 | Configures named console logger, debug level, timestamp format and handler. |
| `get_db_connect_str()` | 47–100 | Reads selected INI section, resolves ODBC driver, creates connection string. |
| `parse_date_value()` | 103–109 | Parses `YYYY-MM-DD` as a Python date; raises parsing error when invalid. |
| `table_full_name()` | 112–113 | Constructs `LIQUIDITY.LST_SF_EXTRACT.<TABLE>`. |
| `get_date_column_for_table()` | 116–119 | Returns `POSDATE` for `NET_LIABILITIES`, otherwise `AS_OF_DATE`. |
| `validate_environment()` | 122–126 | Uppercases and accepts only `DEV`, `UAT`, `PROD`. |
| `check_table_exists()` | 129–142 | Checks schema and table in `INFORMATION_SCHEMA.TABLES`. |
| `get_column_metadata()` | 145–178 | Queries SQL Server catalog and returns column attributes. |
| `determine_insertable_columns()` | 181–206 | Removes computed, identity, and rowversion fields; verifies date column. |
| `build_insert_sql()` | 209–227 | Constructs the same-table `INSERT ... SELECT` query with placeholders. |
| `get_row_count()` | 230–237 | Counts table rows for a supplied date. |
| `ensure_environment_config()` | 240–243 | Checks INI path and invokes configuration/driver validation. |
| `print_preview()` | 246–274 | Prints proposed operation, row counts, generated SQL and parameters. |
| `run_backfill()` | 277–391 | Runs preflight, optional insert, validation, commit/rollback, and cleanup. |
| `build_parser()` | 394–402 | Defines the six command-line options and required argument choices. |
| `main()` | 405–415 | Parses options, delegates to run, handles outer exceptions. |
| `if __name__ == "__main__":` | 418–420 | Changes working directory and exits with main's status. |

## 13. Boundaries, omissions, and points not proven

These are **limitations of what the uploaded code can establish**, not claims that a particular installation is broken.

- **No live source evidence:** The actual config file, installed drivers, servers, permissions, table definitions, row values and run outputs were not supplied.
- **Not an external ETL job:** The shown code does not import source rows from a file, API or different table. It cannot establish where the table's pre-existing rows originally came from.
- **Metadata context may matter:** SQL Server catalog checks use the configured connection database; the insert uses the constant `LIQUIDITY`. The config value is not visible here.
- **Date constraints are limited:** A valid date and different source/target dates are required; chronology and business-day rules are not checked.
- **Schema constraints remain external:** The code excludes certain column categories, but it does not guarantee that all inserts will succeed against every actual SQL Server schema/constraint.
- **No strict post-insert equality:** The minimum-count rule accepts values **greater than** the source count.
- **No content-level reconciliation:** The code contains no row-wise comparison of source and destination, no business-key comparison and no checksum.
- **No explicit concurrency locking:** The Python logic checks that the target is empty, but does not explicitly acquire a target-date lock to exclude other writers. This absence is not evidence of an actual race.
- **Logging is console-based:** The script creates a stream handler and prints to stdout; no log file, scheduler, notification integration or monitoring backend is defined.
- **One run = one approved table/date pair:** There is no built-in iteration over all 12 tables or a date range.
- **Execution enabled still shows the preview first:** The `MODE: PREVIEW` message is printed in both execution modes, before the optional insert starts.

**Evidence:** code sections listed above, primarily lines 12–29, 47–100, 129–235, 246–274, 277–420.

## 14. Plain-English glossary

| Term | Meaning here |
|---|---|
| **Backfill** | Creating records for a requested date by copying existing records for another date. |
| **Source date** | Date used to select which existing records to copy. |
| **Target date** | Date placed into new records. |
| **Metadata** | Information SQL Server reports about the columns in a table. |
| **ODBC / `pyodbc`** | Driver mechanism and Python package used to connect to SQL Server. |
| **Preview** | Reads and checks information, prints proposed SQL; omitting `--EXECUTE` prevents the insert branch. |
| **Transaction** | Database work that can be committed when checks pass or rolled back on caught errors. |
| **Identity column** | SQL Server-generated identity field the script excludes from the copied insert-column list. |
| **Rowversion** | SQL Server `timestamp`-type category excluded from the explicit insert-column list. |

## 15. Evidence register and final conclusion

**Directly examined:** `Pasted text(7).txt`, the entire supplied 420-line Python script. All source-line locations above refer to that file. No external documentation or inaccessible system has been represented as verified evidence.

**Not examined:** INI contents, live database system catalog results, real record counts, stored procedures, deployment configuration, runtime logs, or a witnessed execution.

**Conclusion:** The code implements a one-table, one-date-pair, metadata-driven SQL Server backfill with pre-insert checks, a non-writing default preview, explicit opt-in execution, and a target-row-count threshold before commit. Its exact live effects cannot be certified without runtime and database evidence.
