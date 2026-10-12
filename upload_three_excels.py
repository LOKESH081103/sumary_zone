import os
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine

# Update these paths to your three Excel files.
EXCEL_FILES = [
    Path("asset_tagging_1.xlsx"),
    Path("asset_tagging_2.xlsx"),
    Path("asset_tagging_3.xlsx"),
]

TABLE_NAME = "asset_tagging"


def clean_headers(columns):
    """Normalize headers consistently, including date-like Excel headers."""
    cleaned = []

    for col in columns:
        if hasattr(col, "strftime"):
            name = col.strftime("%B %y")
        elif isinstance(col, str) and any(token in col for token in ("/", "-", "00:00")):
            try:
                parsed = pd.to_datetime(col, dayfirst=True, errors="raise")
                name = parsed.strftime("%B %y")
            except (ValueError, TypeError):
                name = col
        else:
            name = str(col)

        name = name.strip().lower().replace(" ", "_")
        name = pd.Series([name]).str.replace(r"[^a-zA-Z0-9_]", "", regex=True).iloc[0]
        cleaned.append(name)

    # Make repeated header names unique (for example, column, column_1).
    seen = {}
    unique = []
    for name in cleaned:
        count = seen.get(name, 0)
        unique.append(name if count == 0 else f"{name}_{count}")
        seen[name] = count + 1
    return unique


def load_workbook(path):
    """Read every sheet from one workbook and normalize its headers."""
    sheets = pd.read_excel(path, sheet_name=None, header=0)
    frames = []

    for sheet_name, frame in sheets.items():
        frame.columns = clean_headers(frame.columns)
        # Preserve the source workbook and sheet for traceability.
        frame["_source_file"] = path.name
        frame["_source_sheet"] = str(sheet_name)
        frames.append(frame)

    if not frames:
        raise ValueError(f"No worksheets found in {path}")

    return frames


def main():
    missing = [str(path) for path in EXCEL_FILES if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "Could not find these Excel files. Put them beside this script or edit "
            f"EXCEL_FILES in the script: {', '.join(missing)}"
        )

    # Combine all sheets from all three workbooks.
    frames = []
    for path in EXCEL_FILES:
        frames.extend(load_workbook(path))

    # Ensure the same cleaned headers are present in every sheet.
    expected = list(frames[0].columns)
    for number, frame in enumerate(frames[1:], start=2):
        if list(frame.columns) != expected:
            raise ValueError(
                f"Header mismatch in input sheet #{number}.\n"
                f"Expected: {expected}\n"
                f"Found:    {list(frame.columns)}"
            )

    combined = pd.concat(frames, ignore_index=True)

    # Convert likely date/time columns where possible. Invalid values remain null.
    date_keywords = ("date", "month", "year", "time", "period")
    for column in combined.columns:
        if column.startswith("_"):
            continue
        if any(keyword in column for keyword in date_keywords):
            if pd.api.types.is_numeric_dtype(combined[column]):
                combined[column] = pd.to_datetime(
                    combined[column], unit="D", origin="1899-12-30", errors="coerce"
                )
            else:
                combined[column] = pd.to_datetime(
                    combined[column], dayfirst=True, errors="coerce"
                )

    # Set this variable before running; do not hard-code database passwords here.
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise EnvironmentError(
            "DATABASE_URL is not set. Set it to your PostgreSQL connection URL "
            "before running this script."
        )

    engine = create_engine(db_url)
    try:
        # 'replace' recreates the table each run, matching your original script.
        # Use if_exists='append' instead if you intend to add rows on each run.
        combined.to_sql(
            TABLE_NAME,
            engine,
            if_exists="replace",
            index=False,
            chunksize=1000,
            method="multi",
        )
    finally:
        engine.dispose()

    print(
        f"Upload complete: {len(combined):,} rows from {len(EXCEL_FILES)} workbooks "
        f"written to table '{TABLE_NAME}'."
    )


if __name__ == "__main__":
    main()
