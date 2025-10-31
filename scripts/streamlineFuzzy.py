import pandas as pd
import pyodbc
import time
import os
import glob
import re
from sqlalchemy import create_engine
from datetime import datetime
from pathlib import Path
from rapidfuzz import fuzz, process
from difflib import ndiff
from openpyxl import Workbook
from difflib import Differ
from difflib import Differ, SequenceMatcher

# CONSTANTS
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATE_STR = datetime.now().strftime('%m%d%Y')
QUERY = "Query.sql"
OUTPUT = f"QueryResult{DATE_STR}.csv"
DRIVER = "ODBC Driver 17 for SQL Server"
SERVER = "YOUR_SERVER_NAME"
DATABASE = "YOUR_DATABASE_NAME"
CONNECTION = (
    f"Driver={DRIVER};"
    f"Server={SERVER}; "
    f"Database = {DATABASE};"
    "Trusted_Connection=yes;"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)

def main():
    # SQL QUERY
    # Connect to server & database
    print(f"Connecting to {SERVER}, database: {DATABASE}")
    engine = create_engine("mssql+pyodbc://", creator=lambda: pyodbc.connect(CONNECTION))

    # Resolve paths for input query
    sql_path = Path(__file__).resolve().parent.parent /'query' / QUERY
    SQL = sql_path.read_text(encoding="utf-8-sig")

    print(f"Running query {QUERY}...")
    startTime = time.time()

    # Execute query
    with engine.begin() as conn:
        conn.exec_driver_sql("USE [Chartfinder_Snap];")
        df1 = pd.read_sql_query(SQL, conn)
        df1['Zip'] = df1['Zip'].apply(
            lambda x: str(int(x)).zfill(5) 
            if pd.notna(x) and str(x).strip() != '' and str(x).isdigit() and int(x) !=0 and len(str(int(x))) >= 3 
            else x
            )
    
    endTime = time.time()
    executionTime = endTime - startTime
    print(f"Query complete! Total Time: {int(executionTime //60)} minutes {executionTime %60:.1f} seconds.")

    # PRE FORMATTING
    print("Running pre fuzzy formatting...")
    df1.columns = df1.columns.str.strip()

    df1 = df1.rename(columns={"Address 1&2 Concat": "address"})

    # Insert Regional Assignment Column
    pnp_column = df1.columns.get_loc("PNPCode")
    df1.insert(pnp_column + 1, "regional_assignment", "")

    # Assign Regional Owner Formula
    regional_assignment_map = {
        "Breadbasket": "PLACEHOLDER",
        "Pacific": "PLACEHOLDER",
        "Rockies": "PLACEHOLDER",
        "Bridges & Tunnels": "PLACEHOLDER",
        "NY" : "PLACEHOLDER",
        "Upper Midwest": "PLACEHOLDER",
        "FL": "PLACEHOLDER",
        "The Bayou": "PLACEHOLDER",
        "The Rust Belt": "PLACEHOLDER",
        "Allegany": "PLACEHOLDER",
        "Outer Banks": "PLACEHOLDER",
        "Boston & The Beltway": "PLACEHOLDER",
    }

    df1["regional_assignment"] = (df1["Sub Region"].astype(str).str.strip().map(regional_assignment_map).fillna(""))

    # Prepare sheet2
    df2 = df1[["address", 'Zip']].rename(columns={"Zip": "zip_code"})

    # Open ROI Table
    pattern = os.path.join(SCRIPT_DIR, '..', 'ROITable', 'ROITables*.xlsx')
    matches = glob.glob(pattern)
    df_roi= pd.read_excel(matches[0], sheet_name="HealthPortSiteInfo", engine='openpyxl')
    df_roi.columns = df_roi.columns.str.strip()

    address1 = df_roi.get("Address", "").fillna("").astype(str).str.strip()
    address2 = df_roi.get("Address 2", df_roi.get("Address2", "")).fillna("").astype(str).str.strip()
    roi_addresses = (address1 + " " + address2).str.replace(r"\s+", " ", regex=True).str.strip()
    roi_zips = df_roi.get("ZIP", df_roi.get("Zip", "")).fillna("").astype(str).str.strip()

    df2.insert(2, "roi_addresses", "")
    df2.insert(3, "roi_zips", "")
    n = min(len(df2), len(df_roi))
    df2.loc[:n-1, "roi_addresses"] = roi_addresses.iloc[:n].values
    df2.loc[:n-1, "roi_zips"] = roi_zips.iloc[:n].values

    # Zip Normalization
    df2['roi_zips'] = df2['roi_zips'].apply(
        lambda x: str(int(x)).zfill(5) 
        if pd.notna(x) and str(x).strip() != '' and str(x).isdigit() and int(x) !=0 and len(str(int(x))) >= 3 
        else x)
    
    # SHCode
    roi_lookup = pd.read_excel(matches[0], sheet_name="HealthPort_RDO", usecols="A:Y", engine='openpyxl')
    roi_key = roi_lookup.iloc[:, 0].astype(str).str.strip().str.upper()
    roi_value = roi_lookup.iloc[:, 24]
    vlookup_map = dict(zip(roi_key, roi_value))
    roi_keys = df_roi.iloc[:, 0].astype(str).str.strip().str.upper()
    shcode_series = roi_keys.map(vlookup_map).fillna("")
    df2.insert(4, "SHCODE", "")
    n = min(len(df2), len(shcode_series))
    df2.loc[:n-1, "SHCODE"] = shcode_series.iloc[:n].astype(str).values
    print('Fuzzy formatting complete!')

    # FUZZY LOOKUP PROCESS
    print("\nBeginning fuzzy lookup process...")
    df2.info()

    df2.head()

    df1.columns = ( df1.columns
        .str.strip()
        .str.lower()
        .str.replace(' ', '_')
        .str.replace('-', '_')
        .str.replace('(', '')
        .str.replace(')', '')
        .str.replace('?', '')
        .str.replace('\'', '') 
    )
    df1.columns


    df2.columns = (df2.columns
        .str.strip()
        .str.lower()
        .str.replace(' ', '_')
        .str.replace('-', '_')
        .str.replace('(', '')
        .str.replace(')', '')
        .str.replace('?', '')
        .str.replace('\'', '') 
    )
    df2.columns

    # Function to classify differences using difflib
    def classify_difference(original, matched):
        differ = Differ()
        diff = list(differ.compare(original, matched))
        spelling_diff_count = 0
        numerical_diff = False

        for change in diff:
            if change.startswith("- ") or change.startswith("+ "):
                char = change[2:]
                if char.isdigit():
                    numerical_diff = True
                else:
                    spelling_diff_count += 1

        if numerical_diff:
            return "Numerical"
        elif spelling_diff_count == 1:
            return "1 Letter Off"
        elif spelling_diff_count == 2:
            return "2 Letters Off"
        elif spelling_diff_count == 3:
            return "3 Letters Off"
        elif spelling_diff_count > 3:
            return ">3 Letters Off"
        else:
            return "No Difference"

    # Function to highlight differences
    def highlight_differences(original, matched):
        differ = Differ()
        diff = list(differ.compare(original, matched))
        highlighted = []
        for change in diff:
            if change.startswith("- "):
                highlighted.append(f"[{change[2:]}]")
            elif change.startswith("+ "):
                highlighted.append(f"[{change[2:]}]")
            elif change.startswith("  "):
                highlighted.append(change[2:])
        return " ".join(highlighted)

    # Fuzzy match + ZIP check
    def fuzzy_match_with_zip(address, zip_code, roi_addresses, roi_zips, threshold):
        if not isinstance(address, str):
            return None, None, 0, "No Difference", False, ""

        match = process.extractOne(address, roi_addresses, scorer=fuzz.ratio, score_cutoff=threshold)
        if match:
            best_match, score, idx = match[0], match[1], match[2]
            matched_zip = str(roi_zips[idx])
            zip_code_str = str(zip_code)
            zip_match = fuzz.ratio(zip_code_str, matched_zip) > 85
            diff_type = classify_difference(address, best_match)
            differences = highlight_differences(address, best_match)
            return best_match, matched_zip, score, diff_type, zip_match, differences
        return None, None, 0, "No Difference", False, ""

    # Outreach and region lookup
    def lookup_details(address, df1_addresses, df1_outreach_ids, df1_reg_assignments):
        match = process.extractOne(address, df1_addresses, scorer=fuzz.ratio)
        if match:
            idx = df1_addresses.index(match[0])
            return df1_outreach_ids[idx], df1_reg_assignments[idx]
        return None, None

    # Preprocessing
    df1['address'] = df1['address'].astype(str).str.strip().str.lower()
    df2['address'] = df2['address'].astype(str).str.strip().str.lower()
    df2['roi_addresses'] = df2['roi_addresses'].astype(str).str.strip().str.lower()

    # Filter out rows with SH07 or SH09 in SHCODE
    mask_exclude = ~df2['shcode'].str.contains("SH07|SH09", na=False)
    df2_filtered = df2[mask_exclude].copy()

    # Thresholds
    threshold_high = 95
    threshold_medium = 90
    threshold_low = 85

    # Prepare lists
    roi_addresses = df2_filtered['roi_addresses'].tolist()
    roi_zips = df2_filtered['roi_zips'].astype(str).tolist()
    df1_addresses = df1['address'].tolist()
    df1_outreach_ids = df1['outreachid'].astype(str).tolist()
    df1_reg_assignments = df1['regional_assignment'].astype(str).tolist()

    # Match
    startTime = time.time()
    print("Fuzzy matching started...")
    results = df2_filtered.apply(
        lambda row: fuzzy_match_with_zip(row['address'], row['zip_code'], roi_addresses, roi_zips, threshold_low),
        axis=1
    )

    df2_filtered['matched_address'], df2_filtered['matched_zip'], df2_filtered['match_score'], \
    df2_filtered['difference_type'], df2_filtered['zip_match'], df2_filtered['differences'] = zip(*results)

    # Lookup details
    details = df2_filtered['address'].apply(lambda x: lookup_details(x, df1_addresses, df1_outreach_ids, df1_reg_assignments))
    df2_filtered['outreach_id'], df2_filtered['regional_assignment'] = zip(*details)

    # Buckets
    df_95_above = df2_filtered[df2_filtered['match_score'] >= threshold_high].copy()
    df_90_to_94 = df2_filtered[(df2_filtered['match_score'] >= threshold_medium) & (df2_filtered['match_score'] < threshold_high)].copy()
    df_85_to_89 = df2_filtered[(df2_filtered['match_score'] >= threshold_low) & (df2_filtered['match_score'] < threshold_medium)].copy()

    # Drop extras
    columns_to_drop = ['roi_addresses', 'roi_zips', 'zip_match']
    df_95_above.drop(columns=columns_to_drop, inplace=True, errors='ignore')
    df_90_to_94.drop(columns=columns_to_drop, inplace=True, errors='ignore')
    df_85_to_89.drop(columns=columns_to_drop, inplace=True, errors='ignore')

   # ZIP Match Validation
    def add_zip_match(df, left='zip_code', right='matched_zip'):
        df = df.copy()
        def five(z):
            if pd.isna(z): return ""
            m = re.search(r"\b(\d{5})\b", str(z))
            return m.group(1) if m else ""
        z1 = df[left].map(five)
        z2 = df[right].map(five)
        z2 = df["Zip match"] = (z1 == z2)
        return df
    
    df_95_above = add_zip_match(df_95_above, 'zip_code', 'matched_zip')
    df_90_to_94 = add_zip_match(df_90_to_94 , 'zip_code', 'matched_zip')
    df_85_to_89 = add_zip_match(df_85_to_89, 'zip_code', 'matched_zip')

    endTime = time.time()
    executionTime = endTime - startTime
    print(f"Matching process complete! Total execution time: {int(executionTime //60)} minutes {executionTime %60:.1f} seconds")
    
    # Save
    print("\nSaving file...")
    output_fle = f"Fuzzy{DATE_STR}.xlsx"
    with pd.ExcelWriter(output_fle, engine="openpyxl") as writer:
        df_95_above.to_excel(writer, sheet_name="Match_95_and_Above", index=False)
        df_90_to_94.to_excel(writer, sheet_name="Match_90_to_94", index=False)
        df_85_to_89.to_excel(writer, sheet_name="Match_85_to_89", index=False)

    print(f"Final results saved as {output_fle}")


if __name__ == "__main__":
    main()



