import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel to Tabular Converter", page_icon="📊", layout="wide")

def process_excel(file):
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

st.title("📊 Excel to Tabular Converter")
st.markdown("""
Upload a hierarchical Excel report to convert it into a flat, tabular layout.
This application extracts **Account (Tili_ryhmä)** headers and applies them as a new column for their associated transaction rows.
Total rows (`Yhteensä`) and account headers are removed to provide purely tabular data.
""")

uploaded_file = st.file_uploader("Choose an Excel report file", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing data..."):
            tabular_df, summary_df = process_excel(uploaded_file)
            
        if not tabular_df.empty:
            st.success("File processed successfully!")
            
            st.subheader("Data Preview")
            st.dataframe(tabular_df.head(100), use_container_width=True)
            
            if summary_df is not None and not summary_df.empty:
                st.subheader("Cost Summary per Vendor")
                st.dataframe(summary_df.head(100), use_container_width=True)
            
            # Create Excel file in memory for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                tabular_df.to_excel(writer, index=False, sheet_name='Tabular Data')
                if summary_df is not None and not summary_df.empty:
                    summary_df.to_excel(writer, index=False, sheet_name='Summary by Vendor')
            
            st.download_button(
                label="📥 Download Tabular Excel",
                data=buffer.getvalue(),
                file_name="Converted_Tabular_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("No data could be extracted. Please check the structure of the uploaded Excel file or verify if it contains transaction rows.")
            
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
