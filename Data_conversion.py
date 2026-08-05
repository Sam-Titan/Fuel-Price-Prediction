import pandas as pd
import re
from datetime import datetime

FILE_1 = "dataset 1.xls"
FILE_2 = "dataset 2.xlsx"

def parse_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.to_datetime(x).normalize()

    s = str(x).strip()
    if not s:
        return None

    # remove common prefixes
    s = re.sub(r"(?i)^as on\s*", "", s)

    # try a few common formats
    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y", "%d-%m-%y",
        "%d/%m/%Y", "%d/%m/%y",
        "%d.%m.%Y", "%d.%m.%y",
    ]
    for fmt in formats:
        try:
            return pd.to_datetime(datetime.strptime(s, fmt)).normalize()
        except Exception:
            pass

    # generic fallback
    try:
        return pd.to_datetime(s, errors="raise", dayfirst=True).normalize()
    except Exception:
        return None

def to_num(x):
    if pd.isna(x):
        return None
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return None

def is_clean_3col(df):
    if df.shape[1] < 3:
        return False
    first_three = [str(c).strip().lower() if not pd.isna(c) else "" for c in df.iloc[0, :3].tolist()]
    return ("date" in first_three[0]) and ("petrol" in first_three[1]) and ("diesel" in first_three[2])

def extract_clean_3col(df):
    temp = df.iloc[:, :3].copy()
    temp.columns = ["Date", "Petrol Price", "Diesel Price"]
    temp["Date"] = temp["Date"].apply(parse_date)
    temp["Petrol Price"] = temp["Petrol Price"].apply(to_num)
    temp["Diesel Price"] = temp["Diesel Price"].apply(to_num)
    temp = temp.dropna(subset=["Date", "Petrol Price", "Diesel Price"])
    return temp[["Date", "Petrol Price", "Diesel Price"]]

def extract_wide_delhi_sheet(df):
    rows = []

    # Find a row containing city headers
    header_row = None
    delhi_cols = []

    for i in range(len(df)):
        vals = [str(v).strip() if not pd.isna(v) else "" for v in df.iloc[i].tolist()]
        if "Delhi" in vals:
            delhi_cols = [j for j, v in enumerate(vals) if v == "Delhi"]
            if len(delhi_cols) >= 1:
                header_row = i
                break

    if header_row is None:
        return pd.DataFrame(columns=["Date", "Petrol Price", "Diesel Price"])

    # Case A: one Delhi column and rows labelled Petrol/Diesel in first column
    if len(delhi_cols) == 1:
        dcol = delhi_cols[0]

        date_val = None
        for i in range(min(len(df), header_row + 10)):
            for cell in df.iloc[i].tolist():
                s = str(cell).strip() if not pd.isna(cell) else ""
                m = re.search(r"(?i)as on\s*([0-9./-]+)", s)
                if m:
                    date_val = parse_date(m.group(1))
                    if date_val is not None:
                        break
            if date_val is not None:
                break

        petrol = diesel = None
        for i in range(header_row + 1, len(df)):
            label = str(df.iat[i, 0]).strip().lower() if not pd.isna(df.iat[i, 0]) else ""
            if label == "petrol":
                petrol = to_num(df.iat[i, dcol])
            elif label == "diesel":
                diesel = to_num(df.iat[i, dcol])

        if date_val is not None and petrol is not None and diesel is not None:
            rows.append((date_val, petrol, diesel))

    # Case B: two Delhi columns side-by-side or a daily table
    if len(delhi_cols) >= 2:
        pcol, dcol = delhi_cols[0], delhi_cols[1]

        # choose a date column from the same sheet
        date_col = None
        for c in range(df.shape[1]):
            sample = df.iloc[:, c].dropna().head(20).tolist()
            score = sum(1 for v in sample if parse_date(v) is not None)
            if score >= 3:
                date_col = c
                break

        if date_col is None:
            date_col = 0

        for i in range(header_row + 1, len(df)):
            dt = parse_date(df.iat[i, date_col]) if date_col < df.shape[1] else None
            p = to_num(df.iat[i, pcol]) if pcol < df.shape[1] else None
            d = to_num(df.iat[i, dcol]) if dcol < df.shape[1] else None

            if dt is not None and p is not None and d is not None:
                rows.append((dt, p, d))

    out = pd.DataFrame(rows, columns=["Date", "Petrol Price", "Diesel Price"])
    return out

def extract_any_sheet(df):
    # 1) already-clean 3-column format
    if df.shape[1] >= 3 and is_clean_3col(df):
        return extract_clean_3col(df)

    # 2) if first row looks like headers
    if df.shape[0] > 1:
        first_row = [str(v).strip().lower() if not pd.isna(v) else "" for v in df.iloc[0].tolist()]
        if ("date" in first_row[0]) and ("petrol" in " ".join(first_row)) and ("diesel" in " ".join(first_row)):
            return extract_clean_3col(df)

    # 3) Delhi-wide sheet
    wide = extract_wide_delhi_sheet(df)
    if not wide.empty:
        return wide

    # 4) fallback: scan rows that already look like date, petrol, diesel
    rows = []
    for i in range(len(df)):
        vals = df.iloc[i].tolist()
        if len(vals) < 3:
            continue
        dt = parse_date(vals[0])
        p = to_num(vals[1])
        d = to_num(vals[2])
        if dt is not None and p is not None and d is not None:
            rows.append((dt, p, d))

    return pd.DataFrame(rows, columns=["Date", "Petrol Price", "Diesel Price"])

def read_workbook(path):
    if path.lower().endswith(".xls"):
        engine = "xlrd"
    else:
        engine = "openpyxl"

    sheets = pd.read_excel(path, sheet_name=None, header=None, engine=engine)
    all_parts = []

    for sheet_name, df in sheets.items():
        part = extract_any_sheet(df)
        if not part.empty:
            all_parts.append(part)

    if not all_parts:
        return pd.DataFrame(columns=["Date", "Petrol Price", "Diesel Price"])

    result = pd.concat(all_parts, ignore_index=True)
    result["Date"] = pd.to_datetime(result["Date"]).dt.date
    result = (
        result.dropna()
              .drop_duplicates(subset=["Date", "Petrol Price", "Diesel Price"])
              .sort_values("Date")
              .reset_index(drop=True)
    )
    return result

# Read both files
df1 = read_workbook(FILE_1)
df2 = read_workbook(FILE_2)

# Merge and keep only the 3 requested columns
merged = (
    pd.concat([df1, df2], ignore_index=True)
      .drop_duplicates(subset=["Date", "Petrol Price", "Diesel Price"])
      .sort_values("Date")
      .reset_index(drop=True)
)

# Save output
merged.to_csv("delhi_petrol_diesel_merged.csv", index=False)
merged.to_excel("delhi_petrol_diesel_merged.xlsx", index=False)

print(merged.head())
print(merged.tail())
print("Total rows:", len(merged))