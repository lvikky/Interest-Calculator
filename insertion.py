#!/usr/bin/env python3
"""Insert, complete, or delete a reporting-status row in SQL Server.

The script only uses REPORTING_DATE; no second date is needed.
Without --EXECUTE, the script previews the SQL and makes no changes.
Connection settings follow the existing data_backfill(1).py script.
"""

import argparse
import configparser
import datetime as dt
import logging
import os
import re
import sys

import pyodbc

# Replace TABLE_NAME once you know the actual table name.
# Confirm DATABASE_NAME and SCHEMA_NAME before running.
DATABASE_NAME = "LIQUIDITY"
SCHEMA_NAME = "LST_SF_EXTRACT"
TABLE_NAME = "ETL_data_trigger"

DEFAULT_HEAD = "SOURCE_DATA"
DEFAULT_FLAG = True
COMPLETED_STATUS = "completed"

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("reporting-status")


def get_db_connect_str(config_file, environment):
    """Build a SQL Server connection string from [SQLSERVER_<ENV>] in an INI file."""
    config = configparser.ConfigParser()
    if not config.read(config_file):
        raise FileNotFoundError(f"Config file not found: {config_file}")

    section = f"SQLSERVER_{environment}"
    if section not in config:
        raise ValueError(f"Section [{section}] not found in {config_file}")

    db = config[section]
    configured_driver = db["DRIVER"].strip().strip("{}")
    available = {driver.lower(): driver for driver in pyodbc.drivers()}
    driver = available.get(configured_driver.lower())

    if driver is None:
        for fallback in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server"):
            driver = available.get(fallback.lower())
            if driver:
                logger.warning("Configured driver unavailable; using %s", driver)
                break
    if driver is None:
        raise RuntimeError("No supported SQL Server ODBC driver is installed")

    return (
        f"Driver={{{driver}}};"
        f"Server={db['SERVER']};"
        f"Database={db['DATABASE']};"
        f"Trusted_Connection={db.get('Trusted_Connection', 'yes')};"
        f"Encrypt={db.get('Encrypt', 'no')};"
        f"TrustServerCertificate={db.get('TrustServerCertificate', 'yes')};"
    )


def valid_date(value):
    """Require an ISO-formatted reporting date."""
    try:
        date = dt.datetime.strptime(value, "%Y-%m-%d").date()
        if date.isoformat() != value:
            raise ValueError("Noncanonical date")
        return date.isoformat()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("REPORTING_DATE must be YYYY-MM-DD") from exc


def full_table_name():
    """Validate configured identifiers, which cannot use SQL ? placeholders."""
    for label, value in (
        ("DATABASE_NAME", DATABASE_NAME),
        ("SCHEMA_NAME", SCHEMA_NAME),
        ("TABLE_NAME", TABLE_NAME),
    ):
        if value == "REPLACE_WITH_TABLE_NAME" or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value):
            raise ValueError(f"Set a valid {label} in the script: {value}")
    return f"[{DATABASE_NAME}].[{SCHEMA_NAME}].[{TABLE_NAME}]"


def build_parser():
    parser = argparse.ArgumentParser(description="Manage a reporting-status row")
    parser.add_argument("--ENVIRONMENT", choices=("DEV", "UAT", "PROD"), required=True)
    parser.add_argument("--ACTION", choices=("INSERT", "UPDATE", "DELETE"), required=True)
    parser.add_argument("--REPORTING_DATE", type=valid_date, required=True,
                        help="Date of the row to insert, complete, or delete")
    parser.add_argument("--HEAD", default=DEFAULT_HEAD,
                        help="Row identifier (default: SOURCE_DATA)")
    parser.add_argument("--CONFIG_FILE", default="dbconfig.ini")
    parser.add_argument("--EXECUTE", action="store_true",
                        help="Apply and commit the change (default: preview only)")
    return parser


def run(args):
    if not args.HEAD.strip():
        raise ValueError("HEAD cannot be blank")

    table = full_table_name()
    config_file = args.CONFIG_FILE
    if not os.path.isabs(config_file):
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_file)

    conn = pyodbc.connect(get_db_connect_str(config_file, args.ENVIRONMENT),
                          autocommit=False, timeout=120)
    try:
        cursor = conn.cursor()
        try:
            # Verify the expected columns exist before attempting to change any rows.
            cursor.execute(f"SELECT TOP (0) [HEAD], [REPORTING_DATE], [FLAG], [STATUS] FROM {table}")

            # Find the row using HEAD + the one REPORTING_DATE. Lock until rollback/commit.
            cursor.execute(
                f"SELECT COUNT(*) FROM {table} WITH (UPDLOCK, HOLDLOCK) "
                "WHERE [HEAD] = ? AND [REPORTING_DATE] = ?",
                (args.HEAD, args.REPORTING_DATE),
            )
            matches = cursor.fetchone()[0]

            if args.ACTION == "INSERT":
                if matches:
                    raise ValueError("A row already exists for this HEAD and REPORTING_DATE")
                sql = (f"INSERT INTO {table} ([HEAD], [REPORTING_DATE], [FLAG], [STATUS]) "
                       "VALUES (?, ?, ?, ?)")
                # None is sent as a real SQL NULL, not the text 'null'.
                params = (args.HEAD, args.REPORTING_DATE, DEFAULT_FLAG, None)
            else:
                if matches != 1:
                    raise ValueError(f"Expected exactly 1 matching row, found {matches}; no changes made")
                if args.ACTION == "UPDATE":
                    sql = (f"UPDATE {table} SET [STATUS] = ? "
                           "WHERE [HEAD] = ? AND [REPORTING_DATE] = ?")
                    params = (COMPLETED_STATUS, args.HEAD, args.REPORTING_DATE)
                else:  # DELETE
                    sql = f"DELETE FROM {table} WHERE [HEAD] = ? AND [REPORTING_DATE] = ?"
                    params = (args.HEAD, args.REPORTING_DATE)

            print(f"\nMode: {'EXECUTE' if args.EXECUTE else 'PREVIEW (no changes)'}")
            print(f"Environment: {args.ENVIRONMENT} | Action: {args.ACTION} | Table: {table}")
            print(f"HEAD: {args.HEAD} | REPORTING_DATE: {args.REPORTING_DATE} | Matches: {matches}")
            print(f"SQL: {sql}\nParameters: {params}\n")

            if not args.EXECUTE:
                conn.rollback()
                print("Preview complete. No changes made. Add --EXECUTE to commit.")
                return 0

            cursor.execute(sql, params)
            if cursor.rowcount != 1:
                raise RuntimeError(f"Expected 1 affected row, got {cursor.rowcount}")
            conn.commit()
            print(f"Success: {args.ACTION} committed (1 row affected).")
            return 0
        finally:
            cursor.close()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def main():
    args = build_parser().parse_args()
    try:
        return run(args)
    except Exception as exc:
        logger.error("Operation failed: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
