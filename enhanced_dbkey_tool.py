import oracledb
import tkinter as tk
from tkinter import messagebox, scrolledtext
import tkinter.simpledialog as simpledialog
from datetime import datetime
import os
from openpyxl import load_workbook
import pandas as pd
import pickle
import hashlib

# === Oracle Setup ===
oracle_client_path = r"\\whqhome01p\home$\akoreboi\docs\PDF Favorites\Outlook Archive\instantclient_21_10"
conn = None
oracle_available = False

try:
    oracledb.init_oracle_client(lib_dir=oracle_client_path)
    oracle_available = True
except Exception as e:
    print(f"Oracle client initialization warning (running in demo mode): {e}")
    oracle_available = False

def reconnect_db():
    global conn, oracle_available
    if not oracle_available:
        print("Running in demo mode - Oracle database not available")
        conn = None
        return False
    
    try:
        conn = oracledb.connect(
            user="pub",
            password="pub",
            dsn=oracledb.makedsn("wrepdb.sfwmd.gov", 1521, "wrep")
        )
        return True
    except Exception as e:
        print(f"Database connection error: {e}")
        conn = None
        oracle_available = False
        return False

reconnect_db()

try:
    os.environ['TNS_ADMIN'] = r"\\ad.sfwmd.gov\dfsroot\tnsnames"
except Exception as e:
    print(f"TNS_ADMIN setup warning: {e}")

# ======= EDIT MEMORY CLASS - ONLY FOR METADATA SHEET =======
class EditMemory:
    def __init__(self, storage_file="manual_edits_memory.pkl"):
        self.storage_file = storage_file
        self.manual_edits = {}  # {(dbkey, date_str): {'value': val, 'code': code, 'timestamp': time}}
        self.oracle_checksums = {}  # {(dbkey, date_str): hash} to detect Oracle changes
        self.load_edits()
    
    def save_edits(self):
        try:
            with open(self.storage_file, 'wb') as f:
                pickle.dump({
                    'manual_edits': self.manual_edits,
                    'oracle_checksums': self.oracle_checksums
                }, f)
            print(f"Saved {len(self.manual_edits)} manual edits to memory")
        except Exception as e:
            print(f"Failed to save edits: {e}")
    
    def load_edits(self):
        if not os.path.exists(self.storage_file):
            return
        try:
            with open(self.storage_file, 'rb') as f:
                data = pickle.load(f)
                self.manual_edits = data.get('manual_edits', {})
                self.oracle_checksums = data.get('oracle_checksums', {})
            print(f"Loaded {len(self.manual_edits)} manual edits from memory")
        except Exception as e:
            print(f"Failed to load edits: {e}")
            self.manual_edits = {}
            self.oracle_checksums = {}
    
    def record_edit(self, dbkey, date_obj, value, code):
        date_str = date_obj.strftime('%Y-%m-%d') if hasattr(date_obj, 'strftime') else str(date_obj)
        key = (dbkey.upper(), date_str)
        self.manual_edits[key] = {
            'value': value,
            'code': code,
            'timestamp': datetime.now()
        }
        self.save_edits()
    
    def get_edit(self, dbkey, date_obj):
        date_str = date_obj.strftime('%Y-%m-%d') if hasattr(date_obj, 'strftime') else str(date_obj)
        key = (dbkey.upper(), date_str)
        return self.manual_edits.get(key)
    
    def should_use_edit(self, dbkey, date_obj, oracle_value, oracle_code):
        date_str = date_obj.strftime('%Y-%m-%d') if hasattr(date_obj, 'strftime') else str(date_obj)
        key = (dbkey.upper(), date_str)
        
        if key not in self.manual_edits:
            return False
        
        # Check if Oracle data changed since edit
        current_hash = hashlib.md5(f"{oracle_value}|{oracle_code}".encode()).hexdigest()
        
        if key in self.oracle_checksums:
            stored_hash = self.oracle_checksums[key]
            if current_hash != stored_hash:
                # Oracle data changed, remove manual edit and use Oracle
                del self.manual_edits[key]
                self.save_edits()
                print(f"Oracle data changed for {dbkey} on {date_str}, using updated Oracle data")
                return False
        
        # Update checksum and use manual edit
        self.oracle_checksums[key] = current_hash
        return True

# Global edit memory instance
edit_memory = EditMemory()

# ======= ENHANCED FILTERING FUNCTIONS WITH COMPREHENSIVE TABLE DATA =======
def get_all_dbkey_filter_info():
    """Get comprehensive DBKEY information with all columns from the table for filtering"""
    if not oracle_available:
        # Enhanced demo data with all the table columns
        sample_data = {
            'SOURCE_DBKEY': [
                '91490', '15043', '84522', '73295', '68148', '52964', '47852', '39573', '28395', '17527',
                '', '84523', '73296', '', '52965', '47853', '', '28396', '17528', '91491',
                '15044', '', '73297', '68149', '', '47854', '39574', '', '17529', '91492'
            ],
            'PREF_DBKEY': [
                '', '15042', '', '73294', '', '', '47851', '39572', '', '',
                '94783', '', '', '68147', '', '', '35847', '', '18374', '',
                '', '67342', '', '', '49275', '', '', '21693', '', '12965'
            ],
            'STATION_NAME': [
                'S336', 'S80', 'S333', 'S197', 'G76', 'S65', 'L31E', 'L31W', 'S18C', 'G211',
                'S357', 'S12A', 'S12B', 'S12C', 'S12D', 'G72', 'G74', 'G75', 'S175', 'S176',
                'S177', 'S178', 'S179', 'S180', 'S181', 'S182', 'S183', 'S184', 'S185', 'S186'
            ],
            'STAFF_IN_CHARGE': [
                'John Smith', 'Jane Doe', 'Mike Johnson', 'Sarah Wilson', 'David Brown', 'Lisa Garcia', 
                'Robert Miller', 'Jennifer Davis', 'Michael Wilson', 'Amanda Taylor', 'Kevin Anderson', 
                'Maria Rodriguez', 'James Martinez', 'Emily Johnson', 'Christopher Lee', 'Ashley Brown',
                'Matthew Davis', 'Jessica Garcia', 'Daniel Rodriguez', 'Nicole Martinez', 'Solomon Kumar', 
                'Zhu Wei', 'Alex Thompson', 'Rachel Chen', 'Brian Park', 'Samantha Kim', 'Derek Liu', 
                'Melissa Wang', 'Jason Zhou', 'Stephanie Yang'
            ],
            'PREFERENCE': [
                'Monthly', 'Weekly', 'Quarterly', 'Daily', 'Bi-weekly', 'Annual', 'Semi-annual', 'Tri-monthly',
                'Monthly', 'Weekly', 'Quarterly', 'Daily', 'Bi-weekly', 'Annual', 'Semi-annual', 'Tri-monthly',
                'Monthly', 'Weekly', 'Quarterly', 'Daily', 'Bi-weekly', 'Annual', 'Semi-annual', 'Tri-monthly',
                'Monthly', 'Weekly', 'Quarterly', 'Daily', 'Bi-weekly', 'Annual'
            ],
            'LOCATION': [
                'South Florida', 'Central Florida', 'North Florida', 'East Coast', 'West Coast', 'Panhandle',
                'South Florida', 'Central Florida', 'North Florida', 'East Coast', 'West Coast', 'Panhandle',
                'South Florida', 'Central Florida', 'North Florida', 'East Coast', 'West Coast', 'Panhandle',
                'South Florida', 'Central Florida', 'North Florida', 'East Coast', 'West Coast', 'Panhandle',
                'South Florida', 'Central Florida', 'North Florida', 'East Coast', 'West Coast', 'Panhandle'
            ],
            'DISTRICT': [
                'District 1', 'District 2', 'District 3', 'District 4', 'District 5', 'District 1',
                'District 2', 'District 3', 'District 4', 'District 5', 'District 1', 'District 2',
                'District 3', 'District 4', 'District 5', 'District 1', 'District 2', 'District 3',
                'District 4', 'District 5', 'District 1', 'District 2', 'District 3', 'District 4',
                'District 5', 'District 1', 'District 2', 'District 3', 'District 4', 'District 5'
            ],
            'BASIN': [
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress',
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress',
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress',
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress',
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress',
                'Everglades', 'Kissimmee', 'Lower East Coast', 'Upper East Coast', 'Big Cypress'
            ]
        }
        df = pd.DataFrame(sample_data)
        return df
    
    try:
        cursor = conn.cursor()
        # Comprehensive query to get all table columns
        query = """
        SELECT DISTINCT 
            t.SOURCE_DBKEY, 
            t.PREF_DBKEY, 
            t.STATION_NAME,
            t.STAFF_IN_CHARGE, 
            t.PREFERENCE,
            t.LOCATION,
            t.DISTRICT,
            t.BASIN
        FROM STATION_DBKEY_TABLE t
        WHERE (t.SOURCE_DBKEY IS NOT NULL AND t.SOURCE_DBKEY != '')
           OR (t.PREF_DBKEY IS NOT NULL AND t.PREF_DBKEY != '')
        ORDER BY t.STATION_NAME
        """
        cursor.execute(query)
        results = cursor.fetchall()
        
        df = pd.DataFrame(results, columns=[
            'SOURCE_DBKEY', 'PREF_DBKEY', 'STATION_NAME', 'STAFF_IN_CHARGE', 
            'PREFERENCE', 'LOCATION', 'DISTRICT', 'BASIN'
        ])
        return df
    except Exception as e:
        print(f"Database query failed, using demo data: {e}")
        # Return enhanced demo data if query fails
        return pd.DataFrame(sample_data)

def get_unique_values_for_column(column_name):
    """Get unique values for any column from the table"""
    df = get_all_dbkey_filter_info()
    if df.empty or column_name not in df.columns:
        return []
    
    # Handle empty, N/A, TBD values
    unique_values = df[column_name].dropna().unique()
    unique_values = [str(val).strip() for val in unique_values if str(val).strip() not in ['', 'N/A', 'TBD', 'nan']]
    return sorted(unique_values)

def get_filtered_pairs(selected_filters):
    """
    Get filtered station pairs based on multiple column filters
    Returns pairs where source and pref are matched by station but handle missing values properly
    """
    df = get_all_dbkey_filter_info()
    if df.empty:
        return []
    
    # Apply all filters
    filtered_df = df.copy()
    
    for column, selected_values in selected_filters.items():
        if selected_values and len(selected_values) > 0:
            # Handle null/empty values in filtering
            filtered_df = filtered_df[
                (filtered_df[column].isin(selected_values)) |
                (filtered_df[column].isna()) |
                (filtered_df[column].str.strip().isin(['', 'N/A', 'TBD']))
            ]
    
    if filtered_df.empty:
        return []
    
    # Group by station and create pairs
    results = []
    for _, row in filtered_df.iterrows():
        station_name = row['STATION_NAME']
        source_dbkey = row['SOURCE_DBKEY']
        pref_dbkey = row['PREF_DBKEY']
        
        # Handle empty/N/A/TBD/blank values
        source_valid = pd.notna(source_dbkey) and str(source_dbkey).strip() not in ['', 'N/A', 'TBD']
        pref_valid = pd.notna(pref_dbkey) and str(pref_dbkey).strip() not in ['', 'N/A', 'TBD']
        
        # Always include the station pair, even if one side is missing
        if source_valid or pref_valid:
            results.append({
                'station_name': station_name,
                'source_dbkey': source_dbkey if source_valid else None,
                'pref_dbkey': pref_dbkey if pref_valid else None,
                'staff': row['STAFF_IN_CHARGE'],
                'preference': row['PREFERENCE'],
                'location': row['LOCATION'],
                'district': row['DISTRICT'],
                'basin': row['BASIN']
            })
    
    return results

# ======= ALL ORIGINAL FUNCTIONS PRESERVED =======
def load_pref_prn_to_oracle(file_path, batch_size=1000):
    import oracledb
    import pandas as pd
    from tkinter import messagebox

    try:
        conn = oracledb.connect(user="pub", password="pub", dsn="wrep")
        cursor = conn.cursor()

        df = pd.read_csv(file_path, sep='\t', header=None, dtype=str)
        df.columns = ['DBKEY', 'DAILY_DATE', 'VALUE', 'CODE']

        df['DAILY_DATE_ORA'] = pd.to_datetime(
            df['DAILY_DATE'].str.strip().str.upper(),
            format='%d-%b-%y',
            errors='coerce'
        ).dt.strftime('%d-%b-%y').str.upper()

        df = df.dropna(subset=['DAILY_DATE_ORA'])

        sql_insert = """
            INSERT INTO DM_DAILY_DATA (DBKEY, DAILY_DATE, VALUE, CODE)
            VALUES (:dbkey, TO_DATE(:daily_date, 'DD-MON-RR'), :value, :code)
        """

        rows = []
        for _, row in df.iterrows():
            rows.append({
                'dbkey': str(row['DBKEY']).upper(),
                'daily_date': row['DAILY_DATE_ORA'],
                'value': '' if pd.isna(row['VALUE']) else str(row['VALUE']),
                'code': '' if pd.isna(row['CODE']) else str(row['CODE'])
            })

            if len(rows) >= batch_size:
                cursor.executemany(sql_insert, rows)
                conn.commit()
                rows = []

        if rows:
            cursor.executemany(sql_insert, rows)
            conn.commit()

        messagebox.showinfo("Success", f"Loaded PREF data from\n{file_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load PREF file:\n{e}")

    finally:
        cursor.close()
        conn.close()

def get_all_dbkey_info():
    if not oracle_available:
        # Sample data for demo
        sample_data = {
            'DBKEY': ['S80_H', 'S80_T', 'S333_H', 'S333_T', 'PREF_S80_H', 'PREF_S80_T'],
            'STATION': ['S80', 'S80', 'S333', 'S333', 'S80', 'S80'],
            'FREQUENCY': ['DA', 'DA', 'DA', 'DA', 'DA', 'DA']
        }
        df = pd.DataFrame(sample_data)
        df.columns = df.columns.str.upper()
        return df
    
    query = """
        SELECT DBKEY, STATION, FREQUENCY
        FROM KEYWORD_TAB
        WHERE FREQUENCY = 'DA'
          AND DBKEY NOT LIKE '%BK%'
        ORDER BY STATION, DBKEY
    """
    df = pd.read_sql(query, conn)
    df.columns = df.columns.str.upper()
    return df

def get_dbkey_info(dbkeys):
    if not oracle_available:
        # Demo data
        sample_data = {
            'DBKEY': dbkeys,
            'STATION': [f'Station_{dbk[:3]}' for dbk in dbkeys]
        }
        return pd.DataFrame(sample_data)
    
    q = f"""
        SELECT DBKEY, STATION
        FROM KEYWORD_TAB
        WHERE DBKEY IN ({','.join([':dbkey'+str(i) for i in range(len(dbkeys))])})
          AND FREQUENCY = 'DA'
          AND DBKEY NOT LIKE '%BK%'
    """
    params = {f'dbkey{i}': dbk for i, dbk in enumerate(dbkeys)}
    df = pd.read_sql(q, conn, params=params)
    df.columns = df.columns.str.upper()
    return df

def fetch_data(dbkey):
    if not oracle_available:
        # Sample data for demo
        import numpy as np
        dates = pd.date_range(start='2024-01-01', end='2024-01-31', freq='D')
        np.random.seed(hash(dbkey) % 2**32)  # Consistent data per dbkey
        values = np.random.uniform(10.0, 50.0, len(dates))
        codes = ['M' if np.random.random() > 0.9 else '' for _ in range(len(dates))]
        
        sample_data = {
            'DAILY_DATE': dates,
            'VALUE': [f"{val:.2f}" for val in values],
            'CODE': codes
        }
        return pd.DataFrame(sample_data)
    
    query = """
        SELECT DAILY_DATE, VALUE, CODE
        FROM DM_DAILY_DATA
        WHERE DBKEY = :dbkey
        ORDER BY DAILY_DATE
    """
    df = pd.read_sql(query, conn, params={"dbkey": dbkey})
    return df

def get_head_tail_dbkeys_for_station(station_id):
    if not oracle_available:
        # Demo data
        sample_data = {
            'DBKEY': [f'{station_id}_H', f'{station_id}_T'],
            'STATION_DESC': [f'{station_id} Headwater', f'{station_id} Tailwater'],
            'STATION': [station_id, station_id],
            'TYPE': ['Headwater', 'Tailwater']
        }
        return pd.DataFrame(sample_data)
    
    query = """
        SELECT DBKEY, STATION_DESC, STATION,
            CASE 
                WHEN UPPER(STATION_DESC) LIKE '%HEAD%' THEN 'Headwater'
                WHEN UPPER(STATION_DESC) LIKE '%TAIL%' THEN 'Tailwater'
                ELSE 'Unknown'
            END AS TYPE
        FROM KEYWORD_TAB
        WHERE FREQUENCY = 'DA'
          AND DBKEY NOT LIKE '%BK%'
          AND STATION = :station
    """
    return pd.read_sql(query, conn, params={"station": station_id})

def write_block(df, sheet_name, wb, writer, start_date=None, end_date=None):
    if df.empty:
        return

    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    header_format = wb.add_format({'bold': True, 'bg_color': '#D9E1F2'})
    date_format = wb.add_format({'num_format': 'yyyy-mm-dd'})

    all_dates = pd.date_range(
        start=start_date if start_date else df['DAILY_DATE'].min(),
        end=end_date if end_date else datetime.today().date()
    )

    stations = df['Station'].unique()
    col = 1

    for station in stations:
        sub = df[df['Station'] == station]

        source_dbkeys = sub[sub['Type'].str.upper() == 'SOURCE']['DBKEY'].unique()
        pref_dbkeys = sub[sub['Type'].str.upper() == 'PREF']['DBKEY'].unique()

        max_pairs = max(len(source_dbkeys), len(pref_dbkeys))

        for i in range(max_pairs):
            source_dbkey = source_dbkeys[i] if i < len(source_dbkeys) else ''
            pref_dbkey = pref_dbkeys[i] if i < len(pref_dbkeys) else ''

            source_data = sub[(sub['DBKEY'] == source_dbkey) & (sub['Type'].str.upper() == 'SOURCE')][['DAILY_DATE','VALUE','CODE']].copy()
            pref_data = sub[(sub['DBKEY'] == pref_dbkey) & (sub['Type'].str.upper() == 'PREF')][['DAILY_DATE','VALUE','CODE']].copy()

            def align_data(df_subset):
                df_subset = df_subset.drop_duplicates(subset=['DAILY_DATE'], keep='last')
                df_subset = df_subset.set_index('DAILY_DATE').reindex(all_dates).reset_index()
                df_subset['VALUE'] = df_subset['VALUE'].fillna('')
                df_subset['CODE'] = df_subset['CODE'].fillna('')
                return df_subset

            source_data_aligned = align_data(source_data) if not source_data.empty else pd.DataFrame({'index': all_dates, 'VALUE': ['']*len(all_dates), 'CODE': ['']*len(all_dates)})
            pref_data_aligned = align_data(pref_data) if not pref_data.empty else pd.DataFrame({'index': all_dates, 'VALUE': ['']*len(all_dates), 'CODE': ['']*len(all_dates)})

            station_name = ''
            if source_dbkey:
                station_name = df[(df['DBKEY'] == source_dbkey)]['Station'].iloc[0]
            elif pref_dbkey:
                station_name = df[(df['DBKEY'] == pref_dbkey)]['Station'].iloc[0]

            ws.write(0, col, station_name, header_format)

            ws.write(1, col, f"Source: {source_dbkey}", header_format)
            ws.write(1, col + 2, f"PREF: {pref_dbkey}", header_format)

            ws.write(2, col, "VALUE", header_format)
            ws.write(2, col + 1, "CODE", header_format)
            ws.write(2, col + 2, "VALUE", header_format)
            ws.write(2, col + 3, "CODE", header_format)

            for row_num, date in enumerate(all_dates, start=3):
                if col == 1 and i == 0:
                    ws.write_datetime(row_num, 0, date, date_format)

                ws.write(row_num, col, source_data_aligned.iloc[row_num - 3]['VALUE'])
                ws.write(row_num, col + 1, source_data_aligned.iloc[row_num - 3]['CODE'])
                ws.write(row_num, col + 2, pref_data_aligned.iloc[row_num - 3]['VALUE'])
                ws.write(row_num, col + 3, pref_data_aligned.iloc[row_num - 3]['CODE'])

            col += 5

    ws.write(0, 0, "DAILY_DATE", header_format)
    ws.freeze_panes(3, 1)

def load_existing_metadata_from_excel(file_path: str, sheet_name: str = "Metadata") -> pd.DataFrame:
    """ENHANCED: Now records manual edits in memory"""
    if not os.path.exists(file_path):
        return pd.DataFrame()

    wb = load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()
    
    ws = wb[sheet_name]

    data = []
    col = 2

    while col <= ws.max_column:
        source_header_cell = ws.cell(row=2, column=col)
        pref_header_cell = ws.cell(row=2, column=col + 2)

        # Read SOURCE block
        if source_header_cell.value and isinstance(source_header_cell.value, str) and source_header_cell.value.upper().startswith("SOURCE:"):
            source_dbkey = source_header_cell.value.split(":")[1].strip().upper()
            station = ws.cell(row=1, column=col).value or "Unknown"
            row = 4
            while True:
                date_val = ws.cell(row=row, column=1).value
                if date_val is None:
                    break
                val = ws.cell(row=row, column=col).value
                code = ws.cell(row=row, column=col + 1).value
                if val is not None or code is not None:
                    edit_date = pd.to_datetime(date_val)
                    
                    # ENHANCED: Record as manual edit in memory
                    if val or code:
                        edit_memory.record_edit(source_dbkey, edit_date, val, code)
                    
                    data.append({
                        "DAILY_DATE": edit_date,
                        "VALUE": val,
                        "CODE": code,
                        "DBKEY": source_dbkey,
                        "Type": "Source",
                        "Station": station
                    })
                row += 1

        # Read PREF block
        if pref_header_cell.value and isinstance(pref_header_cell.value, str) and pref_header_cell.value.upper().startswith("PREF:"):
            pref_dbkey = pref_header_cell.value.split(":")[1].strip().upper()
            station = ws.cell(row=1, column=col).value or "Unknown"
            row = 4
            while True:
                date_val = ws.cell(row=row, column=1).value
                if date_val is None:
                    break
                val = ws.cell(row=row, column=col + 2).value
                code = ws.cell(row=row, column=col + 3).value
                if val is not None or code is not None:
                    edit_date = pd.to_datetime(date_val)
                    
                    # ENHANCED: Record as manual edit in memory
                    if val or code:
                        edit_memory.record_edit(pref_dbkey, edit_date, val, code)
                    
                    data.append({
                        "DAILY_DATE": edit_date,
                        "VALUE": val,
                        "CODE": code,
                        "DBKEY": pref_dbkey,
                        "Type": "PREF",
                        "Station": station
                    })
                row += 1

        col += 5

    df = pd.DataFrame(data)
    return df

def write_source_pref_excel(pairs, output_file, start_date=None, end_date=None):
    """ENHANCED with edit persistence memory for metadata sheet and improved pairing logic"""
    reconnect_db()

    # Load existing workbook if it exists
    existing_wb = None
    if os.path.exists(output_file):
        existing_wb = load_workbook(output_file, data_only=True)

    # Load existing manual edits from Excel Metadata sheet AND record them in memory
    existing_df = load_existing_metadata_from_excel(output_file)

    # Extract requested DBKEYs (uppercase)
    requested_dbkeys = {dbk.upper() for dbk, _typ in pairs}

    if not existing_df.empty:
        # Filter existing edits by requested DBKEYs
        existing_df = existing_df[existing_df['DBKEY'].str.upper().isin(requested_dbkeys)]

        # Filter by date range if provided
        if start_date:
            existing_df = existing_df[existing_df['DAILY_DATE'] >= start_date]
        if end_date:
            existing_df = existing_df[existing_df['DAILY_DATE'] <= end_date]

        print(f"Loaded existing manual edits filtered by requested DBKEYs: {len(existing_df)} rows")
    else:
        print("No existing manual edits found")

    # Fetch Oracle data for requested DBKEYs & Types
    source_rows = []
    for dbkey, typ in pairs:
        if not dbkey:  # Skip empty/None DBKEYs
            continue
            
        df = fetch_data(dbkey)
        if df.empty:
            continue
        df['DBKEY'] = dbkey
        df['Type'] = typ
        info = get_dbkey_info([dbkey])
        df['Station'] = info.iloc[0]['STATION'] if not info.empty else 'Unknown'
        df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'])

        if start_date:
            df = df[df['DAILY_DATE'] >= start_date]
        if end_date:
            df = df[df['DAILY_DATE'] <= end_date]

        if df.empty:
            continue

        # ENHANCED: Apply edit memory to Oracle data
        for idx, row in df.iterrows():
            oracle_val = row['VALUE']
            oracle_code = row['CODE']
            
            # Check if we should use manual edit instead of Oracle data
            if edit_memory.should_use_edit(dbkey, row['DAILY_DATE'], oracle_val, oracle_code):
                edit_data = edit_memory.get_edit(dbkey, row['DAILY_DATE'])
                if edit_data:
                    df.at[idx, 'VALUE'] = edit_data['value']
                    df.at[idx, 'CODE'] = edit_data['code']

        source_rows.append(df)

    oracle_df = pd.concat(source_rows, ignore_index=True) if source_rows else pd.DataFrame()

    # Combine manual edits (existing_df) with Oracle data (oracle_df)
    if not existing_df.empty and not oracle_df.empty:
        oracle_only = oracle_df.merge(
            existing_df[['DBKEY', 'DAILY_DATE']],
            on=['DBKEY', 'DAILY_DATE'],
            how='left',
            indicator=True
        )
        oracle_only = oracle_only[oracle_only['_merge'] == 'left_only'].drop(columns=['_merge'])
        
        # Add manual-only edits (dates that don't exist in Oracle)
        manual_only = existing_df.merge(
            oracle_df[['DBKEY', 'DAILY_DATE']],
            on=['DBKEY', 'DAILY_DATE'],
            how='left',
            indicator=True
        )
        manual_only = manual_only[manual_only['_merge'] == 'left_only'].drop(columns=['_merge'])
        
        combined_df = pd.concat([oracle_df, manual_only], ignore_index=True)
    elif not existing_df.empty:
        combined_df = existing_df.copy()
    else:
        combined_df = oracle_df.copy()

    combined_df['DAILY_DATE'] = pd.to_datetime(combined_df['DAILY_DATE'], errors='coerce')
    combined_df.sort_values(['DBKEY', 'DAILY_DATE'], inplace=True)
    combined_df.drop_duplicates(subset=["DBKEY", "DAILY_DATE"], keep="last", inplace=True)

    # Create Excel writer
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='yyyy-mm-dd')
    wb = writer.book

    # Copy existing sheets except those to overwrite (SourcePrefData, HeadTailData, Metadata)
    sheets_to_overwrite = {"SourcePrefData", "HeadTailData", "Metadata", "Comparison"}
    if existing_wb:
        for sheetname in existing_wb.sheetnames:
            if sheetname not in sheets_to_overwrite:
                source_ws = existing_wb[sheetname]
                target_ws = wb.add_worksheet(sheetname)
                for row in source_ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            if isinstance(cell.value, (datetime, pd.Timestamp)):
                                target_ws.write_datetime(cell.row - 1, cell.column - 1, cell.value)
                            elif isinstance(cell.value, (int, float)):
                                target_ws.write_number(cell.row - 1, cell.column - 1, cell.value)
                            else:
                                target_ws.write(cell.row - 1, cell.column - 1, str(cell.value))

    # Write SourcePrefData sheet (PURE ORACLE DB DATA ONLY - NO EDITS)
    write_pure_oracle_sourceprefdata_sheet(oracle_df, wb, writer, start_date, end_date)

    # Write HeadTailData sheet
    all_stations = oracle_df['Station'].unique() if not oracle_df.empty else []
    ht_rows = []
    for station in all_stations:
        station_id = station
        ht_keys = get_head_tail_dbkeys_for_station(station_id)
        for _, row in ht_keys.iterrows():
            dbkey = row['DBKEY']
            typ = row['TYPE']
            df = fetch_data(dbkey)
            if df.empty:
                continue
            df['DBKEY'] = dbkey
            df['Type'] = typ
            df['Station'] = station
            df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'])

            if start_date:
                df = df[df['DAILY_DATE'] >= start_date]
            if end_date:
                df = df[df['DAILY_DATE'] <= end_date]

            ht_rows.append(df)
    ht_df = pd.concat(ht_rows, ignore_index=True) if ht_rows else pd.DataFrame()

    if not ht_df.empty:
        df_grouped = ht_df.groupby('DBKEY')
        col = 0
        ws = wb.add_worksheet("HeadTailData")
        writer.sheets["HeadTailData"] = ws
        header_format = wb.add_format({'bold': True, 'bg_color': '#D9E1F2'})
        date_format = wb.add_format({'num_format': 'yyyy-mm-dd'})

        for dbkey, group in df_grouped:
            group = group.sort_values('DAILY_DATE')
            station = group['Station'].iloc[0]
            typ = group['Type'].iloc[0]
            title = f"{typ}: {dbkey} ({station})"

            ws.write(0, col, title, header_format)
            ws.write(1, col, 'DAILY_DATE', header_format)
            ws.write(1, col + 1, 'VALUE', header_format)
            ws.write(1, col + 2, 'CODE', header_format)

            for row_num, (_, row) in enumerate(group.iterrows(), start=2):
                ws.write_datetime(row_num, col, row['DAILY_DATE'], date_format)
                ws.write_number(row_num, col + 1, 0 if pd.isna(row['VALUE']) else row['VALUE'])
                ws.write_string(row_num, col + 2, '' if pd.isna(row['CODE']) else str(row['CODE']))

            col += 4

        ws.freeze_panes(2, 0)

    # Write Metadata sheet with EDIT MEMORY INTEGRATION
    metadata_df = combined_df[combined_df['Type'].str.upper().isin(['PREF', 'SOURCE'])].copy()
    write_block(metadata_df, "Metadata", wb, writer, start_date, end_date)

    # Write new Comparison sheet (Metadata format + Difference + Percentage Error)
    source_df = combined_df[combined_df['Type'].str.upper() == 'SOURCE'].copy()
    pref_df = combined_df[combined_df['Type'].str.upper() == 'PREF'].copy()

    write_comparison_sheet(source_df, pref_df, wb, writer, start_date, end_date)

    writer.close()
    
    print(f"Excel file created with edit memory: {output_file}")
    print(f"Manual edits in memory: {len(edit_memory.manual_edits)}")

def write_comparison_sheet(source_df, pref_df, wb, writer, start_date, end_date):
    """
    Creates a comparison sheet with EXACT SAME LAYOUT as metadata sheet PLUS:
    - Difference: Source-PREF (automated calculation)
    - PercentageError: (Source-pref/pref)*100 (automated calculation)
    """
    # Combine both dataframes
    combined_df = pd.concat([source_df, pref_df], ignore_index=True)
    
    if combined_df.empty:
        return

    sheet_name = "Comparison"
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    header_format = wb.add_format({'bold': True, 'bg_color': "#2C3038"})
    calc_format = wb.add_format({'bold': True, 'bg_color': '#FFE599'})  # Yellow for calculations
    date_format = wb.add_format({'num_format': 'yyyy-mm-dd'})
    number_format = wb.add_format({'num_format': '0.000'})

    all_dates = pd.date_range(
        start=start_date if start_date else combined_df['DAILY_DATE'].min(),
        end=end_date if end_date else datetime.today().date()
    )

    stations = combined_df['Station'].unique()
    col = 1

    for station in stations:
        sub = combined_df[combined_df['Station'] == station]

        source_dbkeys = sub[sub['Type'].str.upper() == 'SOURCE']['DBKEY'].unique()
        pref_dbkeys = sub[sub['Type'].str.upper() == 'PREF']['DBKEY'].unique()

        max_pairs = max(len(source_dbkeys), len(pref_dbkeys))

        for i in range(max_pairs):
            source_dbkey = source_dbkeys[i] if i < len(source_dbkeys) else ''
            pref_dbkey = pref_dbkeys[i] if i < len(pref_dbkeys) else ''

            source_data = sub[(sub['DBKEY'] == source_dbkey) & (sub['Type'].str.upper() == 'SOURCE')][['DAILY_DATE','VALUE','CODE']].copy()
            pref_data = sub[(sub['DBKEY'] == pref_dbkey) & (sub['Type'].str.upper() == 'PREF')][['DAILY_DATE','VALUE','CODE']].copy()

            def align_data(df_subset):
                if df_subset.empty:
                    return pd.DataFrame({'index': all_dates, 'VALUE': ['']*len(all_dates), 'CODE': ['']*len(all_dates)})
                df_subset = df_subset.drop_duplicates(subset=['DAILY_DATE'], keep='last')
                df_subset = df_subset.set_index('DAILY_DATE').reindex(all_dates).reset_index()
                df_subset['VALUE'] = df_subset['VALUE'].fillna('')
                df_subset['CODE'] = df_subset['CODE'].fillna('')
                return df_subset

            source_data_aligned = align_data(source_data)
            pref_data_aligned = align_data(pref_data)

            station_name = ''
            if source_dbkey:
                station_name = combined_df[(combined_df['DBKEY'] == source_dbkey)]['Station'].iloc[0]
            elif pref_dbkey:
                station_name = combined_df[(combined_df['DBKEY'] == pref_dbkey)]['Station'].iloc[0]

            # Station header (spans all 7 columns: Source(2) + PREF(2) + Difference(1) + PercentError(1) + Space(1))
            ws.merge_range(0, col, 0, col + 6, station_name, header_format)

            # DBKEY headers
            ws.write(1, col, f"Source: {source_dbkey}", header_format)
            ws.write(1, col + 2, f"PREF: {pref_dbkey}", header_format)
            ws.write(1, col + 4, "Difference: Source-PREF", calc_format)
            ws.write(1, col + 5, "PercentageError: Source-pref/pref*100", calc_format)

            # Column headers
            ws.write(2, col, "VALUE", header_format)
            ws.write(2, col + 1, "CODE", header_format)
            ws.write(2, col + 2, "VALUE", header_format)
            ws.write(2, col + 3, "CODE", header_format)
            ws.write(2, col + 4, "DIFFERENCE", calc_format)
            ws.write(2, col + 5, "PERCENT_ERROR", calc_format)

            # Data rows with automated calculations
            for row_num, date in enumerate(all_dates, start=3):
                if col == 1 and i == 0:
                    ws.write_datetime(row_num, 0, date, date_format)

                # Get source and pref values
                source_val_str = source_data_aligned.iloc[row_num - 3]['VALUE']
                source_code = source_data_aligned.iloc[row_num - 3]['CODE']
                pref_val_str = pref_data_aligned.iloc[row_num - 3]['VALUE']
                pref_code = pref_data_aligned.iloc[row_num - 3]['CODE']

                # Write source data
                ws.write(row_num, col, source_val_str)
                ws.write(row_num, col + 1, source_code)
                
                # Write pref data
                ws.write(row_num, col + 2, pref_val_str)
                ws.write(row_num, col + 3, pref_code)

                # AUTOMATED CALCULATIONS
                try:
                    source_num = None
                    pref_num = None
                    
                    # Convert to numbers for calculations
                    if source_val_str and str(source_val_str).strip() != '':
                        try:
                            source_num = float(str(source_val_str).strip())
                        except:
                            source_num = None
                    
                    if pref_val_str and str(pref_val_str).strip() != '':
                        try:
                            pref_num = float(str(pref_val_str).strip())
                        except:
                            pref_num = None

                    # Calculate DIFFERENCE: Source - PREF
                    if source_num is not None and pref_num is not None:
                        difference = source_num - pref_num
                        ws.write_number(row_num, col + 4, difference, number_format)
                    else:
                        ws.write(row_num, col + 4, "")

                    # Calculate PERCENTAGE ERROR: (Source-pref)/pref*100
                    if source_num is not None and pref_num is not None and pref_num != 0:
                        percent_error = ((source_num - pref_num) / pref_num) * 100
                        ws.write_number(row_num, col + 5, percent_error, number_format)
                    else:
                        ws.write(row_num, col + 5, "")

                except Exception as e:
                    # If calculation fails, leave cells empty
                    ws.write(row_num, col + 4, "")
                    ws.write(row_num, col + 5, "")

            col += 7  # Move to next station (4 data columns + 2 calculation columns + 1 space)

    ws.write(0, 0, "DAILY_DATE", header_format)
    ws.freeze_panes(3, 1)

def write_pure_oracle_sourceprefdata_sheet(oracle_df, wb, writer, start_date, end_date):
    """
    Write SourcePrefData sheet with PURE ORACLE DATA ONLY - NO EDITS APPLIED
    Completely bypasses edit memory system to ensure pure Oracle database data
    """
    # Filter Oracle data by date range only - NO EDIT MEMORY INVOLVEMENT
    pure_oracle_df = oracle_df.copy()
    if start_date:
        pure_oracle_df = pure_oracle_df[pure_oracle_df['DAILY_DATE'] >= start_date]
    if end_date:
        pure_oracle_df = pure_oracle_df[pure_oracle_df['DAILY_DATE'] <= end_date]
    
    if pure_oracle_df.empty:
        return
    
    # Get unique pairs and process them
    pairs_data = {}
    for _, row in pure_oracle_df.iterrows():
        dbkey = row['DBKEY']
        typ = row['Type']
        station = row['Station']
        
        # Group by station for pairing
        if station not in pairs_data:
            pairs_data[station] = {'SOURCE': None, 'PREF': None}
        
        if typ.upper() == 'SOURCE':
            pairs_data[station]['SOURCE'] = dbkey
        elif typ.upper() == 'PREF':
            pairs_data[station]['PREF'] = dbkey
    
    # Create worksheet
    ws = wb.add_worksheet("SourcePrefData")
    writer.sheets["SourcePrefData"] = ws
    
    # Formats
    header_format = wb.add_format({'bold': True, 'bg_color': '#D9E1F2'})
    date_format = wb.add_format({'num_format': 'yyyy-mm-dd'})
    
    # Get all unique dates and sort them
    all_dates = sorted(pure_oracle_df['DAILY_DATE'].unique())
    
    # Write date column
    ws.write(0, 0, "DAILY_DATE", header_format)
    for row_num, date in enumerate(all_dates, start=1):
        ws.write_datetime(row_num, 0, date, date_format)
    
    # Process each station pair
    col = 1
    for station, dbkeys in pairs_data.items():
        source_dbkey = dbkeys['SOURCE']
        pref_dbkey = dbkeys['PREF']
        
        if source_dbkey or pref_dbkey:
            # Write station header
            ws.write(0, col, station, header_format)
            
            # Write SOURCE data if available
            if source_dbkey:
                ws.write(1, col, f"SOURCE: {source_dbkey}", header_format)
                ws.write(1, col + 1, "CODE", header_format)
                
                source_data = pure_oracle_df[(pure_oracle_df['DBKEY'] == source_dbkey) & 
                                           (pure_oracle_df['Type'].str.upper() == 'SOURCE')]
                
                for row_num, date in enumerate(all_dates, start=2):
                    date_data = source_data[source_data['DAILY_DATE'] == date]
                    if not date_data.empty:
                        val = date_data['VALUE'].iloc[0]
                        code = date_data['CODE'].iloc[0]
                        if pd.notna(val):
                            try:
                                ws.write_number(row_num, col, float(val))
                            except:
                                ws.write(row_num, col, str(val))
                        if pd.notna(code):
                            ws.write_string(row_num, col + 1, str(code))
            
            col += 2
            
            # Write PREF data if available
            if pref_dbkey:
                ws.write(1, col, f"PREF: {pref_dbkey}", header_format)
                ws.write(1, col + 1, "CODE", header_format)
                
                pref_data = pure_oracle_df[(pure_oracle_df['DBKEY'] == pref_dbkey) & 
                                         (pure_oracle_df['Type'].str.upper() == 'PREF')]
                
                for row_num, date in enumerate(all_dates, start=2):
                    date_data = pref_data[pref_data['DAILY_DATE'] == date]
                    if not date_data.empty:
                        val = date_data['VALUE'].iloc[0]
                        code = date_data['CODE'].iloc[0]
                        if pd.notna(val):
                            try:
                                ws.write_number(row_num, col, float(val))
                            except:
                                ws.write(row_num, col, str(val))
                        if pd.notna(code):
                            ws.write_string(row_num, col + 1, str(code))
            
            col += 3  # Space between station pairs
    
    # Freeze panes
    ws.freeze_panes(2, 1)

# ======= ALL ORIGINAL FUNCTION BUTTONS PRESERVED =======
def generate_pref_txt_files_from_metadata(file_path="DBKEY_Stations_Output.xlsx"):
    wb = load_workbook(file_path, data_only=True)
    if "Metadata" not in wb.sheetnames:
        messagebox.showerror("Error", "Metadata sheet not found in the Excel file.")
        return
    
    ws = wb["Metadata"]

    saved_files = []
    col = 1
    while col <= ws.max_column:
        header_cell = ws.cell(row=2, column=col)
        if header_cell.value and isinstance(header_cell.value, str) and header_cell.value.upper().startswith("PREF:"):
            try:
                dbkey = header_cell.value.split(":")[1].split("(")[0].strip().upper()
            except Exception:
                col += 1
                continue

            data = {'DBKEY': [], 'DAILY_DATE': [], 'VALUE': [], 'CODE': []}
            row = 4
            while True:
                date_val = ws.cell(row=row, column=1).value
                if date_val is None:
                    break
                val = ws.cell(row=row, column=col).value
                code = ws.cell(row=row, column=col + 1).value

                if val is None and code is None:
                    row += 1
                    continue

                data['DBKEY'].append(dbkey)
                data['DAILY_DATE'].append(date_val)
                data['VALUE'].append(val)
                data['CODE'].append(code)
                row += 1

            if not data['DAILY_DATE']:
                col += 5
                continue

            df = pd.DataFrame(data)
            df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'], errors='coerce').dt.strftime('%d-%b-%y').str.upper()
            df = df[['DBKEY', 'DAILY_DATE', 'VALUE', 'CODE']]

            output_filename = f"{dbkey}.pref.txt"
            df.to_csv(output_filename, sep='\t', index=False, header=False)
            saved_files.append(output_filename)

        col += 5

    messagebox.showinfo("PREF Files Generated", "Exported files:\n" + "\n".join(saved_files))

def generate_pref_upload_file(file_path="DBKEY_Stations_Output.xlsx"):
    wb = load_workbook(file_path, data_only=True)
    if "Metadata" not in wb.sheetnames:
        messagebox.showerror("Error", "Metadata sheet not found in the Excel file.")
        return
    
    ws = wb["Metadata"]

    all_data = {'DBKEY': [], 'DAILY_DATE': [], 'VALUE': [], 'CODE': []}
    col = 1
    while col <= ws.max_column:
        header_cell = ws.cell(row=2, column=col)
        if header_cell.value and isinstance(header_cell.value, str) and header_cell.value.upper().startswith("PREF:"):
            try:
                dbkey = header_cell.value.split(":")[1].split("(")[0].strip().upper()
            except Exception:
                col += 1
                continue

            row = 4
            while True:
                date_val = ws.cell(row=row, column=1).value
                if date_val is None:
                    break
                val = ws.cell(row=row, column=col).value
                code = ws.cell(row=row, column=col + 1).value

                if val is None and code is None:
                    row += 1
                    continue

                all_data['DBKEY'].append(dbkey)
                all_data['DAILY_DATE'].append(date_val)
                all_data['VALUE'].append(val)
                all_data['CODE'].append(code)
                row += 1

        col += 5

    if not all_data['DAILY_DATE']:
        messagebox.showinfo("No Data", "No PREF data found.")
        return

    df = pd.DataFrame(all_data)
    df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'], errors='coerce').dt.strftime('%d-%b-%y').str.upper()
    df = df[['DBKEY', 'DAILY_DATE', 'VALUE', 'CODE']]

    output_filename = "all_pref_upload_data.txt"
    df.to_csv(output_filename, sep='\t', index=False, header=False)

    messagebox.showinfo("Upload File Generated", f"All PREF data exported to:\n{output_filename}")

def update_pref_data_from_excel(file_path="DBKEY_Stations_Output.xlsx"):
    messagebox.showinfo("Update Complete", "PREF data has been reloaded from the Metadata sheet.")

def generate_pref_prn_files_from_metadata(file_path="DBKEY_Stations_Output.xlsx"):
    """FIXED: Create Unix text files for ALL PREF DBKEYs - works like single save but for all"""
    wb = load_workbook(file_path, data_only=True)
    if "Metadata" not in wb.sheetnames:
        messagebox.showerror("Error", "Metadata sheet not found in the Excel file.")
        return
    
    ws = wb["Metadata"]

    saved_files = []
    col = 1
    while col <= ws.max_column:
        header_cell = ws.cell(row=2, column=col)
        if header_cell.value and isinstance(header_cell.value, str) and header_cell.value.upper().startswith("PREF:"):
            try:
                # Extract DBKEY from header (e.g., "PREF: ABC123 (Station)")
                dbkey = header_cell.value.split(":")[1].split("(")[0].strip().upper()
            except Exception:
                col += 1
                continue

            # Collect all data for this PREF DBKEY (same logic as single save)
            data = {'DBKEY': [], 'DAILY_DATE': [], 'VALUE': [], 'CODE': []}
            row = 4  # Start from row 4 (after headers)
            
            while True:
                date_val = ws.cell(row=row, column=1).value
                if date_val is None:
                    break
                    
                val = ws.cell(row=row, column=col).value
                code = ws.cell(row=row, column=col + 1).value

                # Include rows with data (same logic as single save function)
                if val is None and code is None:
                    row += 1
                    continue

                data['DBKEY'].append(dbkey)
                data['DAILY_DATE'].append(date_val)
                data['VALUE'].append(val)
                data['CODE'].append(code)
                row += 1

            # Only create file if we have data
            if data['DAILY_DATE']:
                df = pd.DataFrame(data)
                # Format date as DD-MMM-YY (Oracle format)
                df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'], errors='coerce').dt.strftime('%d-%b-%y').str.upper()
                df = df[['DBKEY', 'DAILY_DATE', 'VALUE', 'CODE']]

                # Save as Unix text file (.prn)
                output_filename = f"{dbkey}.pref.prn"
                df.to_csv(output_filename, sep='\t', index=False, header=False)
                saved_files.append(output_filename)
                print(f"Generated Unix file for PREF DBKEY: {dbkey} with {len(df)} records")

            # Move to next PREF block - same increment logic as single save
            col += 5
        else:
            col += 1

    if saved_files:
        messagebox.showinfo("Unix PREF Files Generated", f"Exported {len(saved_files)} Unix files based on PREF DBKEYs:\n" + "\n".join(saved_files))
    else:
        messagebox.showinfo("No PREF Data", "No PREF DBKEYs found in Metadata sheet to export.")

def save_single_pref_prn_from_metadata(file_path="DBKEY_Stations_Output.xlsx"):
    root = tk.Tk()
    root.withdraw()
    dbkey_input = simpledialog.askstring("Input", "Enter DBKEY to export as .pref.prn file (e.g. ABC123):")
    root.destroy()

    if not dbkey_input:
        messagebox.showinfo("Cancelled", "No DBKEY entered.")
        return

    dbkey_input = dbkey_input.strip().upper()

    wb = load_workbook(file_path, data_only=True)
    if "Metadata" not in wb.sheetnames:
        messagebox.showerror("Error", "Metadata sheet not found in the Excel file.")
        return
    
    ws = wb["Metadata"]

    col = 1
    found_block = None
    while col <= ws.max_column:
        header_cell = ws.cell(row=2, column=col)
        if header_cell.value and isinstance(header_cell.value, str) and header_cell.value.upper().startswith("PREF:"):
            try:
                dbkey = header_cell.value.split(":")[1].split("(")[0].strip().upper()
                if dbkey == dbkey_input:
                    found_block = col
                    break
                col += 5
            except Exception:
                col += 1
        else:
            col += 1

    if not found_block:
        messagebox.showerror("Not Found", f"DBKEY '{dbkey_input}' not found in Metadata sheet.")
        return

    data = {'DBKEY': [], 'DAILY_DATE': [], 'VALUE': [], 'CODE': []}
    row = 4
    while True:
        date_val = ws.cell(row=row, column=1).value
        if date_val is None:
            break
        val = ws.cell(row=row, column=found_block).value
        code = ws.cell(row=row, column=found_block + 1).value

        if val is None and code is None:
            row += 1
            continue

        data['DBKEY'].append(dbkey_input)
        data['DAILY_DATE'].append(date_val)
        data['VALUE'].append(val)
        data['CODE'].append(code)
        row += 1

    if not data['DAILY_DATE']:
        messagebox.showinfo("No Data", f"No data found for DBKEY '{dbkey_input}'.")
        return

    df = pd.DataFrame(data)
    df['DAILY_DATE'] = pd.to_datetime(df['DAILY_DATE'], errors='coerce').dt.strftime('%d-%b-%y').str.upper()
    df = df[['DBKEY', 'DAILY_DATE', 'VALUE', 'CODE']]

    output_filename = f"{dbkey_input}.pref.prn"
    df.to_csv(output_filename, sep='\t', index=False, header=False)

    messagebox.showinfo("File Saved", f"Saved single PREF file:\n{output_filename}")

def load_pref_file_button():
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select .pref.prn file to load to Oracle",
        filetypes=[("PRN files", "*.prn"), ("All files", "*.*")]
    )
    root.destroy()
    
    if file_path:
        load_pref_prn_to_oracle(file_path)

def open_all_dbkeys_window():
    all_dbkeys_df = get_all_dbkey_info()
    win = tk.Toplevel()
    win.title("All DBKEYs (DA only)")
    st = scrolledtext.ScrolledText(win, width=100, height=30)
    st.insert(tk.END, all_dbkeys_df[['DBKEY', 'STATION']].to_string(index=False))
    st.pack(padx=10, pady=10)

# ======= ENHANCED GUI WITH COMPREHENSIVE FILTERING =======
def main_gui():
    root = tk.Tk()
    root.title("Enhanced DBKEY Excel Tool - All Columns Filter")
    root.geometry("1200x800")

    # === NEW COMPREHENSIVE FILTERING SECTION ===
    filter_frame = tk.LabelFrame(root, text="Filter DBKEYs by All Table Columns (Excel-like Filtering)", relief=tk.RIDGE, bd=2, font=('Arial', 12, 'bold'))
    filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # Create scrollable frame for filters
    canvas = tk.Canvas(filter_frame)
    scrollbar = tk.Scrollbar(filter_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # All column filters
    filter_widgets = {}
    column_names = ['STATION_NAME', 'STAFF_IN_CHARGE', 'PREFERENCE', 'LOCATION', 'DISTRICT', 'BASIN']
    
    # Create filter interface for each column
    for i, column in enumerate(column_names):
        col_frame = tk.LabelFrame(scrollable_frame, text=f"Filter by {column.replace('_', ' ').title()}", relief=tk.RIDGE, bd=1)
        col_frame.grid(row=i//3, column=i%3, padx=5, pady=5, sticky="nsew")
        
        listbox = tk.Listbox(col_frame, selectmode=tk.MULTIPLE, height=6, width=25)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        scrollbar_col = tk.Scrollbar(col_frame, orient="vertical", command=listbox.yview)
        scrollbar_col.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.configure(yscrollcommand=scrollbar_col.set)
        
        filter_widgets[column] = {
            'listbox': listbox,
            'frame': col_frame
        }
    
    scrollable_frame.columnconfigure(0, weight=1)
    scrollable_frame.columnconfigure(1, weight=1)
    scrollable_frame.columnconfigure(2, weight=1)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Filter control panel
    control_frame = tk.Frame(filter_frame)
    control_frame.pack(fill=tk.X, padx=5, pady=5)

    load_filters_btn = tk.Button(control_frame, text="Load All Filter Options", width=20, bg="lightblue")
    load_filters_btn.pack(side=tk.LEFT, padx=5)

    apply_filter_btn = tk.Button(control_frame, text="Apply Filters & Generate Pairs", width=25, bg="lightgreen")
    apply_filter_btn.pack(side=tk.LEFT, padx=5)

    clear_all_btn = tk.Button(control_frame, text="Clear All Filters", width=20, bg="lightcoral")
    clear_all_btn.pack(side=tk.LEFT, padx=5)

    filter_status = tk.Label(control_frame, text="Ready to load filter options...", fg="blue")
    filter_status.pack(side=tk.LEFT, padx=10)

    # Results display
    results_frame = tk.LabelFrame(filter_frame, text="Filter Results", relief=tk.RIDGE, bd=1)
    results_frame.pack(fill=tk.X, padx=5, pady=5)

    results_text = tk.Text(results_frame, height=8, width=120, wrap=tk.WORD)
    results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
    
    results_scrollbar = tk.Scrollbar(results_frame, orient="vertical", command=results_text.yview)
    results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    results_text.configure(yscrollcommand=results_scrollbar.set)

    # === SEPARATOR ===
    separator = tk.Frame(root, height=2, bg="gray")
    separator.pack(fill=tk.X, padx=10, pady=5)

    # === ORIGINAL MANUAL INPUT SECTION ===
    manual_frame = tk.LabelFrame(root, text="Manual DBKEY Input", relief=tk.RIDGE, bd=2, font=('Arial', 10, 'bold'))
    manual_frame.pack(fill=tk.X, padx=10, pady=5)

    # --- PREF Textbox ---
    tk.Label(manual_frame, text="Enter PREF DBKEYs (comma separated):").pack(anchor=tk.W, padx=5)
    pref_entry = tk.Entry(manual_frame, width=90)
    pref_entry.pack(padx=10, pady=5)

    # --- SOURCE Textbox ---
    tk.Label(manual_frame, text="Enter SOURCE DBKEYs (comma separated):").pack(anchor=tk.W, padx=5)
    source_entry = tk.Entry(manual_frame, width=90)
    source_entry.pack(padx=10, pady=5)

    # === DATE SECTION ===
    date_frame = tk.LabelFrame(root, text="Date Range (Optional)", relief=tk.RIDGE, bd=2)
    date_frame.pack(fill=tk.X, padx=10, pady=5)

    date_controls = tk.Frame(date_frame)
    date_controls.pack(fill=tk.X, padx=5, pady=5)

    tk.Label(date_controls, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, sticky=tk.W, padx=5)
    start_entry = tk.Entry(date_controls, width=20)
    start_entry.grid(row=0, column=1, padx=5)

    tk.Label(date_controls, text="End Date (YYYY-MM-DD):").grid(row=0, column=2, sticky=tk.W, padx=5)
    end_entry = tk.Entry(date_controls, width=20)
    end_entry.grid(row=0, column=3, padx=5)

    # === FILTERING FUNCTIONALITY ===
    def load_all_filter_options():
        """Load all column filter options into listboxes"""
        try:
            filter_status.config(text="Loading all filter options...", fg="blue")
            
            for column in column_names:
                unique_values = get_unique_values_for_column(column)
                listbox = filter_widgets[column]['listbox']
                listbox.delete(0, tk.END)
                for value in unique_values:
                    listbox.insert(tk.END, value)
            
            filter_status.config(text="All filter options loaded successfully", fg="green")
        except Exception as e:
            filter_status.config(text=f"Error loading filter options: {str(e)}", fg="red")

    def clear_all_filters():
        """Clear all filter selections"""
        for column in column_names:
            listbox = filter_widgets[column]['listbox']
            listbox.selection_clear(0, tk.END)
        results_text.delete(1.0, tk.END)
        filter_status.config(text="All filters cleared", fg="blue")

    def apply_all_filters():
        """Apply all selected filters and generate station pairs with proper source/pref pairing"""
        try:
            filter_status.config(text="Applying filters...", fg="blue")
            
            # Get all selected filters
            selected_filters = {}
            for column in column_names:
                listbox = filter_widgets[column]['listbox']
                selected_indices = listbox.curselection()
                selected_values = [listbox.get(i) for i in selected_indices]
                if selected_values:
                    selected_filters[column] = selected_values
            
            # Get filtered pairs
            filtered_pairs = get_filtered_pairs(selected_filters)
            
            # Display results
            results_text.delete(1.0, tk.END)
            
            if not filtered_pairs:
                results_text.insert(tk.END, "No station pairs found matching the selected filters.\n")
                filter_status.config(text="No pairs found matching filters", fg="orange")
                return
            
            results_text.insert(tk.END, f"Found {len(filtered_pairs)} station pairs matching your filters:\n\n")
            
            # Collect DBKEYs for textbox population
            source_dbkeys = []
            pref_dbkeys = []
            
            for i, pair in enumerate(filtered_pairs, 1):
                station_name = pair['station_name']
                source_dbkey = pair['source_dbkey']
                pref_dbkey = pair['pref_dbkey']
                staff = pair['staff']
                preference = pair['preference']
                location = pair['location']
                district = pair['district']
                basin = pair['basin']
                
                results_text.insert(tk.END, f"{i}. Station: {station_name}\n")
                
                if source_dbkey:
                    results_text.insert(tk.END, f"   SOURCE DBKEY: {source_dbkey}\n")
                    source_dbkeys.append(source_dbkey)
                else:
                    results_text.insert(tk.END, f"   SOURCE DBKEY: [No data available]\n")
                
                if pref_dbkey:
                    results_text.insert(tk.END, f"   PREF DBKEY: {pref_dbkey}\n")
                    pref_dbkeys.append(pref_dbkey)
                else:
                    results_text.insert(tk.END, f"   PREF DBKEY: [No data available]\n")
                
                results_text.insert(tk.END, f"   Staff: {staff} | Preference: {preference}\n")
                results_text.insert(tk.END, f"   Location: {location} | District: {district} | Basin: {basin}\n\n")
            
            # Add filtered DBKEYs to textboxes (append to existing content)
            current_source = source_entry.get().strip()
            current_pref = pref_entry.get().strip()
            
            # Add filtered source DBKEYs
            if source_dbkeys:
                new_source_dbkeys = []
                if current_source:
                    new_source_dbkeys.extend([k.strip() for k in current_source.split(',') if k.strip()])
                new_source_dbkeys.extend(source_dbkeys)
                # Remove duplicates while preserving order
                unique_source = []
                for dbkey in new_source_dbkeys:
                    if dbkey not in unique_source:
                        unique_source.append(dbkey)
                source_entry.delete(0, tk.END)
                source_entry.insert(0, ', '.join(unique_source))
            
            # Add filtered pref DBKEYs
            if pref_dbkeys:
                new_pref_dbkeys = []
                if current_pref:
                    new_pref_dbkeys.extend([k.strip() for k in current_pref.split(',') if k.strip()])
                new_pref_dbkeys.extend(pref_dbkeys)
                # Remove duplicates while preserving order
                unique_pref = []
                for dbkey in new_pref_dbkeys:
                    if dbkey not in unique_pref:
                        unique_pref.append(dbkey)
                pref_entry.delete(0, tk.END)
                pref_entry.insert(0, ', '.join(unique_pref))
            
            filter_status.config(text=f"Applied filters - {len(filtered_pairs)} station pairs found", fg="green")
            
        except Exception as e:
            filter_status.config(text=f"Error applying filters: {str(e)}", fg="red")
            results_text.delete(1.0, tk.END)
            results_text.insert(tk.END, f"Error applying filters: {str(e)}")

    # Bind button commands
    load_filters_btn.config(command=load_all_filter_options)
    apply_filter_btn.config(command=apply_all_filters)
    clear_all_btn.config(command=clear_all_filters)
    
    # Load filter options on startup
    root.after(100, load_all_filter_options)

    def on_generate():
        pref_input = pref_entry.get().strip()
        source_input = source_entry.get().strip()

        # Combine both input boxes into pairs
        pairs = []
        if pref_input:
            for dbk in pref_input.split(','):
                dbk = dbk.strip()
                if dbk:
                    pairs.append((dbk.upper(), 'PREF'))

        if source_input:
            for dbk in source_input.split(','):
                dbk = dbk.strip()
                if dbk:
                    pairs.append((dbk.upper(), 'Source'))

        if not pairs:
            messagebox.showerror("Input Error", "Please enter at least one DBKEY in either textbox.")
            return

        # Start/end date validation
        start_date = None
        end_date = None
        if start_entry.get().strip():
            try:
                start_date = datetime.strptime(start_entry.get().strip(), '%Y-%m-%d')
            except ValueError:
                messagebox.showerror("Date Error", "Start Date format must be YYYY-MM-DD")
                return
        if end_entry.get().strip():
            try:
                end_date = datetime.strptime(end_entry.get().strip(), '%Y-%m-%d')
            except ValueError:
                messagebox.showerror("Date Error", "End Date format must be YYYY-MM-DD")
                return

        output_file = os.path.abspath("DBKEY_Stations_Output.xlsx")
        try:
            write_source_pref_excel(pairs, output_file, start_date, end_date)
            messagebox.showinfo("Success", f"Excel file generated with edit memory:\n{output_file}\n\nYour manual edits will now persist when you change date ranges!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Excel file:\n{e}")

    # === ALL ORIGINAL BUTTONS SECTION ===
    buttons_frame = tk.LabelFrame(root, text="Actions", relief=tk.RIDGE, bd=2, font=('Arial', 10, 'bold'))
    buttons_frame.pack(fill=tk.X, padx=10, pady=5)

    # Original buttons in organized layout
    button_grid = tk.Frame(buttons_frame)
    button_grid.pack(padx=5, pady=5)

    # Row 1
    tk.Button(button_grid, text="Generate Excel", command=on_generate, width=25, bg="lightgreen").grid(row=0, column=0, padx=2, pady=2)
    tk.Button(button_grid, text="Generate PREF TXT Files", command=generate_pref_txt_files_from_metadata, width=25).grid(row=0, column=1, padx=2, pady=2)
    tk.Button(button_grid, text="Generate PREF Upload TXT", command=generate_pref_upload_file, width=25).grid(row=0, column=2, padx=2, pady=2)

    # Row 2
    tk.Button(button_grid, text="Reload Metadata Sheet", command=update_pref_data_from_excel, width=25).grid(row=1, column=0, padx=2, pady=2)
    tk.Button(button_grid, text="Create Unix Files for PREF Load", command=generate_pref_prn_files_from_metadata, width=25).grid(row=1, column=1, padx=2, pady=2)
    tk.Button(button_grid, text="Save Single PREF as .prn", command=save_single_pref_prn_from_metadata, width=25).grid(row=1, column=2, padx=2, pady=2)

    # Row 3
    tk.Button(button_grid, text="Load PREF .prn to Oracle DB", command=load_pref_file_button, width=25).grid(row=2, column=0, padx=2, pady=2)
    tk.Button(button_grid, text="Show All DBKEYs", command=open_all_dbkeys_window, width=25).grid(row=2, column=1, padx=2, pady=2)

    root.mainloop()
    
if __name__ == "__main__":
    main_gui()