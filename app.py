import streamlit as st
import pandas as pd
import openpyxl
import io

# --- Page Configuration ---
st.set_page_config(page_title="Tally VLOOKUP Updater", layout="centered", page_icon="🔗")
st.title("🔗 Master Ledger VLOOKUP Updater")
st.markdown("Upload your **Master Sheet** and your **Segregated Summary Sheet**. This app will match the item names and automatically VLOOKUP the calculated totals into your master sheet.")

# --- File Uploaders ---
st.subheader("1. Upload Files")
master_file = st.file_uploader("Upload Master Sheet (e.g., final_StkSum.xlsx)", type=['xlsx'])
summary_file = st.file_uploader("Upload Segregated Summary Sheet (.xlsx)", type=['xlsx'])

# --- Process Button ---
if st.button("Run VLOOKUP & Update Master"):
    if not master_file or not summary_file:
        st.error("⚠️ Please upload both the Master Sheet and the Summary Sheet.")
    else:
        with st.spinner("Running VLOOKUP mapping..."):
            try:
                # 1. Read the Summary File into a pandas DataFrame
                df_summary = pd.read_excel(summary_file)
                
                # Make sure the summary file has the expected columns. 
                # We assume the first column is 'Item Name' or similar.
                # We will create a dictionary for fast VLOOKUP mapping.
                item_col = df_summary.columns[0]
                
                # Convert summary data into a dictionary for VLOOKUP
                # Adjust column names based on your actual summary file headers
                # We are looking for columns that contain these keywords:
                summary_data = {}
                for idx, row in df_summary.iterrows():
                    item_name = str(row[item_col]).strip()
                    
                    if pd.notna(item_name) and item_name != 'nan':
                        # Default values
                        prod, purch, cons, sales, godown = 0, 0, 0, 0, 0
                        
                        # Flexibly find columns in the summary sheet
                        for col_name in df_summary.columns:
                            col_lower = str(col_name).lower()
                            val = row[col_name]
                            if pd.isna(val): val = 0
                            
                            if 'production' in col_lower or 'prod' in col_lower:
                                prod = val
                            elif 'purchase' in col_lower:
                                purch = val
                            elif 'consumption' in col_lower or 'cons' in col_lower:
                                cons = val
                            elif 'sale' in col_lower:
                                sales = val
                            elif 'godown' in col_lower or 'transfer' in col_lower:
                                godown = val
                                
                        summary_data[item_name] = {
                            'prod': prod,
                            'purch': purch,
                            'cons': cons,
                            'sales': sales,
                            'godown': godown
                        }

                # 2. Open and Update the Master Sheet using openpyxl
                master_file.seek(0)
                wb = openpyxl.load_workbook(master_file)
                ws = wb.active
                
                updates = 0
                
                # Start updating from row 12 (assuming rows 1-11 are headers/spares)
                for row in range(12, ws.max_row + 1):
                    cell_val = ws.cell(row=row, column=1).value
                    
                    if cell_val and isinstance(cell_val, str):
                        master_item_name = cell_val.strip()
                        
                        # --- THE VLOOKUP LOGIC ---
                        if master_item_name in summary_data:
                            m = summary_data[master_item_name]
                            
                            # Clear existing data in target columns to prevent ghost data
                            for col in [18, 19, 20, 21, 24]: # Columns R, S, T, U, X
                                ws.cell(row=row, column=col).value = None
                                
                            # Write new VLOOKUP values if they are greater than 0
                            if m['prod'] > 0: ws.cell(row=row, column=18).value = m['prod']
                            if m['purch'] > 0: ws.cell(row=row, column=19).value = m['purch']
                            if m['cons'] > 0: ws.cell(row=row, column=20).value = m['cons']
                            if m['sales'] > 0: ws.cell(row=row, column=21).value = m['sales']
                            if m['godown'] > 0: ws.cell(row=row, column=24).value = m['godown']
                                
                            updates += 1

                # 3. Save Output
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"✅ VLOOKUP Complete! Successfully mapped and updated {updates} items in the master sheet.")
                
                st.download_button(
                    label="📥 Download VLOOKUP Master Sheet",
                    data=output,
                    file_name="Master_Sheet_VLOOKUP_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.info("Make sure your Summary file has clear column headers like 'Inward Production', 'Outward Sales', etc.")
                
