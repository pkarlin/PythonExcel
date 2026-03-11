import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel to Tabular Converter", page_icon="📊", layout="wide")

def process_ledger(file):
    # Read without header to find the actual header row
    df_raw = pd.read_excel(file, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        # Check if the row contains typical header column names
        row_strs = [str(x).lower().strip() for x in row.values if pd.notna(x)]
        # We also check for 'debet' or 'kredit' to be absolutely sure we found the right header
        if any('tili/päiväys' in s or 'tosite' in s or 'nimi' in s for s in row_strs) and any('debet' in s or 'kredit' in s for s in row_strs):
            header_idx = idx
            break
            
    # Read with correct header, using the detected header index
    df = pd.read_excel(file, header=header_idx)
    
    if df.empty:
        return pd.DataFrame(), None
        
    processed_data = []
    current_account = "Unknown Account"
    
    # Store the column names to verify structure
    col_names = list(df.columns)
    
    if len(col_names) < 2:
        st.error("The uploaded Excel file doesn't have enough columns to be processed properly.")
        return pd.DataFrame(), None
        
    # Find key columns we will use to identify transaction rows
    debet_col = next((c for c in df.columns if 'debet' in str(c).lower()), None)
    kredit_col = next((c for c in df.columns if 'kredit' in str(c).lower()), None)
    tosite_col = next((c for c in df.columns if str(c).lower().strip() == 'tosite'), None)
        
    for index, row in df.iterrows():
        # Clean col1 (usually 'Tili/Päiväys') and col2 (usually 'Nimi/Tositelaji')
        col1_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col2_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
        # We should also check the whole row as a string to find 'Yhteensä' because it might be merged
        row_text_lower = " ".join([str(x).lower().strip() for x in row.values if pd.notna(x)])
        
        # Skip completely empty rows
        if not row_text_lower:
            continue
            
        # Skip Total (Yhteensä) and Transfer (Siirto) rows
        if col1_val.lower().startswith("yhteensä") or col1_val.lower().startswith("siirtoa") or "yhteensä" in row_text_lower:
            continue
            
        # Determine if it's an account grouping row or a transaction row
        is_transaction = False
        
        if debet_col and kredit_col:
            d_val = row[debet_col]
            k_val = row[kredit_col]
            has_d = pd.notna(d_val) and str(d_val).strip() != ""
            has_k = pd.notna(k_val) and str(k_val).strip() != ""
            
            # If a row has either a Debet or Kredit amount, it's a real transaction
            if has_d or has_k:
                is_transaction = True
        elif tosite_col:
            t_val = row[tosite_col]
            if pd.notna(t_val) and str(t_val).strip() != "":
                is_transaction = True
        else:
            # Fallback logic if headers were renamed or missing
            if col1_val and col2_val and len(row) > 2:
                col3_val = row.iloc[2]
                if pd.notna(col3_val) and str(col3_val).strip() != "":
                    is_transaction = True
                    
        if not is_transaction:
            # It's an Account Header row (e.g. "7640" in col1, "Atk-laite..." in col2)
            if col1_val and col2_val:
                current_account = f"{col1_val} {col2_val}"
            elif col1_val:
                current_account = col1_val
            elif col2_val:
                current_account = col2_val
        else:
            # It's a Transaction row
            row_dict = row.to_dict()
            # Apply the extracted account for this transaction
            row_dict['Tili_ryhmä'] = current_account
            processed_data.append(row_dict)
            
    # Create the flattened DataFrame
    result_df = pd.DataFrame(processed_data)
    
    # Reorder columns to place the extracted 'Tili_ryhmä' group first
    if not result_df.empty:
        # Extract Vendor Name from Selite if it exists
        selite_col = next((c for c in result_df.columns if 'selite' in str(c).lower()), None)
        if selite_col:
            # Assumes format "vendorname,invoicenumber"
            # Split by comma and take the first part, strip whitespace. If no comma, take the whole string.
            result_df['Vendor Name'] = result_df[selite_col].astype(str).apply(
                lambda x: x.split(',')[0].strip() if pd.notna(x) and str(x).strip() != "" else ""
            )
            
        # Ensure Tili_ryhmä is first, then Vendor Name if it was created
        first_cols = ['Tili_ryhmä']
        if 'Vendor Name' in result_df.columns:
            first_cols.append('Vendor Name')
            
        cols = first_cols + [col for col in result_df.columns if col not in first_cols]
        result_df = result_df[cols]
        
        # Clean up unneeded completely empty or unnamed columns
        unnamed_cols = [c for c in result_df.columns if str(c).startswith('Unnamed:')]
        for c in unnamed_cols:
            if result_df[c].isna().all() or (result_df[c] == '').all():
                result_df = result_df.drop(columns=[c])
                
    # Calculate Cost and generate Summary
    summary_df = None
    if debet_col and kredit_col:
        # Ensure debet and kredit are numeric
        result_df[debet_col] = pd.to_numeric(result_df[debet_col], errors='coerce').fillna(0)
        result_df[kredit_col] = pd.to_numeric(result_df[kredit_col], errors='coerce').fillna(0)
        
        # Calculate real cost per row
        result_df['Cost'] = result_df[debet_col] - result_df[kredit_col]
        
        # Calculate Non-VAT Cost (Finland VAT 25.5%)
        # Note: round to 2 decimal places for cleaner currency values
        result_df['Non-VAT Cost'] = (result_df['Cost'] / 1.255).round(2)
        
        # Group by Vendor Name if available
        if 'Vendor Name' in result_df.columns:
            # Exclude empty vendor names from summary if any
            summary_base = result_df[result_df['Vendor Name'] != ""]
            
            # Group by Vendor Name and take sum of costs, and the first value of Tili_ryhmä for context
            summary_df = summary_base.groupby('Vendor Name', as_index=False).agg(
                **{
                    'Tili_ryhmä': ('Tili_ryhmä', 'first'),
                    'Non-VAT Cost': ('Non-VAT Cost', 'sum'),
                    'Cost': ('Cost', 'sum'),
                    debet_col: (debet_col, 'sum'),
                    kredit_col: (kredit_col, 'sum')
                }
            )
            # Sort by highest Non-VAT Cost
            summary_df = summary_df.sort_values('Non-VAT Cost', ascending=False)
                
    return result_df, summary_df

def process_ageing(file):
    # Read without header to find the actual header row
    df_raw = pd.read_excel(file, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        # Check if the row contains typical header column names for Ageing report
        row_strs = [str(x).lower().strip() for x in row.values if pd.notna(x)]
        if any('alle 14vrk' in s or 'saldo' in s for s in row_strs):
            header_idx = idx
            break
            
    df = pd.read_excel(file, header=header_idx)
    
    blocks = []
    current_block = []
    current_property = "Unknown Property"
    
    for index, row in df.iterrows():
        col1_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col2_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
        row_text_lower = " ".join([str(x).lower().strip() for x in row.values if pd.notna(x)])
        
        if not row_text_lower:
            continue
            
        has_numbers = False
        # Check from column 2 onwards for numbers (index 2+)
        if len(row) > 2:
            for val in row.values[2:]:
                if pd.notna(val) and str(val).strip() != "":
                    cleaned_val = str(val).strip().replace(',', '.').replace(' ', '').replace('€', '').replace('\xa0', '')
                    try:
                        float(cleaned_val)
                        has_numbers = True
                        break
                    except ValueError:
                        pass
                        
        cleaned_col1 = col1_val.lower().replace(" ", "").replace("\xa0", "")
        
        # Always skip "Sopimus päättynyt" rows (and any row where col1 starts with "sopimus")
        if "sopimus" in cleaned_col1:
            continue
        
        if not has_numbers:
            # Any text-only row finishes an apartment block
            if current_block:
                blocks.append((current_property, current_block))
                current_block = []
                
            # If it looks like a property header, update current_property.
            # A valid property header must start with a digit (property code like 091291001...)
            # to avoid short continuation lines (e.g. "M", "Espoo") being mistaken as a new property.
            if col1_val and not col2_val:
                if "sopimus" in cleaned_col1 or "yhteensä" in cleaned_col1 or "saldo" in cleaned_col1 or cleaned_col1 in ["kmp", "kmhp"]:
                    continue
                if col1_val[0].isdigit():
                    current_property = col1_val
                # else: it's a continuation/overflow line — ignore it
            continue
            
        # If we reach here, the row HAS numbers.
        if "yhteensä" in cleaned_col1 or "saldo" in cleaned_col1 or cleaned_col1 in ["kmp", "kmhp"]:
            # This is the end-of-apartment total row
            if current_block:
                blocks.append((current_property, current_block))
                current_block = []
            continue
            
        # Accumulate all lines (charges, sub-tenant totals) into the current block
        current_block.append(row)
        
    if current_block:
        blocks.append((current_property, current_block))

    processed_data = []
    
    for prop, b in blocks:
        if not b: continue
        
        apt_name = str(b[0].iloc[0]).strip() if pd.notna(b[0].iloc[0]) else ""
        tenant_name = ""
        
        # Look ahead in the block for the tenant name
        # The tenant name usually appears on the 2nd (or 3rd) line in col1
        for _row in b[1:]:
            val = str(_row.iloc[0]).strip() if pd.notna(_row.iloc[0]) else ""
            cleaned_val = val.lower().replace(" ", "").replace("\xa0", "")
            if val and "yhteensä" not in cleaned_val and cleaned_val not in ["kmp", "kmhp"]:
                tenant_name = val
                break
        
        # If the apartment name is empty but tenant_name has a value, it means only one
        # identifying row existed (e.g. a parking unit line), so promote it to Huoneisto.
        if not apt_name and tenant_name:
            apt_name = tenant_name
            tenant_name = ""
                
        # Now walk all rows and keep only those with actual charge types (Selite)
        for _row in b:
            c2_val = str(_row.iloc[1]).strip() if pd.notna(_row.iloc[1]) else ""
            if c2_val: 
                r_dict = _row.to_dict()
                r_dict['Kiinteistö'] = prop
                r_dict['Huoneisto'] = apt_name
                r_dict['Asukas'] = tenant_name
                processed_data.append(r_dict)
            
    result_df = pd.DataFrame(processed_data)
    
    if not result_df.empty:
        col_names = list(result_df.columns)
        first_col = col_names[0]
        second_col = col_names[1]
        
        result_df = result_df.rename(columns={second_col: 'Selite'})
        
        cols = ['Kiinteistö', 'Huoneisto', 'Asukas', 'Selite']
        other_cols = [c for c in result_df.columns if c not in cols and c != first_col]
        result_df = result_df[cols + other_cols]
        
        unnamed_cols = [c for c in result_df.columns if str(c).startswith('Unnamed:')]
        for c in unnamed_cols:
            if result_df[c].isna().all() or (result_df[c] == '').all():
                result_df = result_df.drop(columns=[c])
                
    return result_df, None

def display_results(tabular_df, summary_df, default_filename):
    if not tabular_df.empty:
        st.success("File processed successfully!")
        
        st.subheader("Data Preview")
        st.dataframe(tabular_df.head(100), use_container_width=True)
        
        if summary_df is not None and not summary_df.empty:
            st.subheader("Cost Summary")
            st.dataframe(summary_df.head(100), use_container_width=True)
        
        # Create Excel file in memory for download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            tabular_df.to_excel(writer, index=False, sheet_name='Tabular Data')
            if summary_df is not None and not summary_df.empty:
                summary_df.to_excel(writer, index=False, sheet_name='Summary')
        
        st.download_button(
            label=f"📥 Download {default_filename}",
            data=buffer.getvalue(),
            file_name=default_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.warning("No data could be extracted. Please check the structure of the uploaded Excel file.")

def process_cost_centers(file):
    # Read without header to find the actual header row
    df_raw = pd.read_excel(file, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        # Check if the row contains typical header column names for Cost Centers report
        row_strs = [str(x).lower().strip() for x in row.values if pd.notna(x)]
        if any('alle 14vrk' in s or 'saldo' in s for s in row_strs):
            header_idx = idx
            break
            
    df = pd.read_excel(file, header=header_idx)
    
    if df.empty:
        return pd.DataFrame(), None
        
    processed_data = []
    
    for index, row in df.iterrows():
        col1_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        
        row_text_lower = " ".join([str(x).lower().strip() for x in row.values if pd.notna(x)])
        if not row_text_lower:
            continue
            
        has_numbers = False
        # Check from column 1 onwards for numbers
        if len(row) > 1:
            for val in row.values[1:]:
                if pd.notna(val) and str(val).strip() != "":
                    # handle possible formatted strings
                    stripped = str(val).strip().replace(',', '.').replace(' ', '').replace('€', '').replace('\xa0', '')
                    try:
                        float(stripped)
                        has_numbers = True
                        break
                    except ValueError:
                        pass
                        
        if has_numbers and col1_val:
            if "yhteensä" in col1_val.lower() or "saldo" == col1_val.lower():
                continue # Skip total row
                
            row_dict = row.to_dict()
            processed_data.append(row_dict)
            
    result_df = pd.DataFrame(processed_data)
    
    if not result_df.empty:
        first_col = result_df.columns[0]
        # Rename the first column to indicate Cost Center if it doesn't already have a good name
        if str(first_col).startswith('Unnamed:'):
            result_df = result_df.rename(columns={first_col: 'Kustannuspaikka'})
        
        # Clean up unnamed empty columns
        unnamed_cols = [c for c in result_df.columns if str(c).startswith('Unnamed:')]
        for c in unnamed_cols:
            if result_df[c].isna().all() or (result_df[c] == '').all():
                result_df = result_df.drop(columns=[c])
                
    return result_df, None

st.title("📊 Excel to Tabular Converter")
st.markdown("Choose the correct bucket below to upload your Excel file.")

col1, col2, col3 = st.columns(3)

with col1:
    st.header("1. General Ledger")
    st.markdown("""
    Convert **Transactions (Pääkirja)** reports. Extracts **Account (Tili_ryhmä)** headers.
    """)
    uploaded_ledger = st.file_uploader("Upload General Ledger report", type=["xls", "xlsx"], key="ledger")
    
    if uploaded_ledger is not None:
        try:
            with st.spinner("Processing General Ledger..."):
                tabular_df, summary_df = process_ledger(uploaded_ledger)
            display_results(tabular_df, summary_df, "Converted_General_Ledger.xlsx")
        except Exception as e:
            st.error(f"Error processing General Ledger: {e}")

with col2:
    st.header("2. Ageing Report")
    st.markdown("""
    Convert **Ageing (Saamisten ikäjakauma)** reports. Extracts **Property (Kiinteistö)** headers.
    """)
    uploaded_ageing = st.file_uploader("Upload Ageing report", type=["xls", "xlsx"], key="ageing")
    
    if uploaded_ageing is not None:
        try:
            with st.spinner("Processing Ageing Report..."):
                tabular_df, summary_df = process_ageing(uploaded_ageing)
            display_results(tabular_df, summary_df, "Converted_Ageing_Report.xlsx")
        except Exception as e:
            st.error(f"Error processing Ageing Report: {e}")

with col3:
    st.header("3. Cost Centers")
    st.markdown("""
    Convert **Cost Centers (Saamiset kustannuspaikoittain)** reports. Flattens data rows.
    """)
    uploaded_cost = st.file_uploader("Upload Cost Centers report", type=["xls", "xlsx"], key="cost_centers")
    
    if uploaded_cost is not None:
        try:
            with st.spinner("Processing Cost Centers Report..."):
                tabular_df, summary_df = process_cost_centers(uploaded_cost)
            display_results(tabular_df, summary_df, "Converted_Cost_Centers_Report.xlsx")
        except Exception as e:
            st.error(f"Error processing Cost Centers Report: {e}")
