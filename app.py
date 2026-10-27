import streamlit as st
import pandas as pd
import openpyxl
import math
import io

# --- Page Configuration ---
st.set_page_config(page_title="Tally Ledger Updater", layout="centered", page_icon="📊")
st.title("📊 Tally Stock Ledger Updater")
st.markdown("Upload your Master Sheet and your raw Tally exports to automatically calculate and update the ledger based on the **original summary rules**.")

def clean_val(val):
    """Helper function to clean and convert string numbers to floats."""
    try:
        if pd.isna(val): return 0.0
        v = float(str(val).replace(',', '').strip())
        return v if not math.isnan(v) else 0.0
    except:
        return 0.0

# --- File Uploaders ---
st.subheader("1. Upload Files")
master_file = st.file_uploader("Upload Master Sheet (e.g., final_StkSum.xlsx)", type=['xlsx'])
tally_files = st.file_uploader("Upload Raw Tally Files (.xls or .xlsx)", type=['xls', 'xlsx'], accept_multiple_files=True)

# --- Process Button ---
if st.button("Process & Generate Updated Ledger"):
    if not master_file or not tally_files:
        st.error("⚠️ Please upload both the master sheet and at least one Tally file.")
    else:
        with st.spinner("Processing files and extracting data..."):
            metrics_dict = {}
            
            # --- Step 1: Parse Raw Tally Files ---
            for f in tally_files:
                try:
                    # Handle both standard Excel and Tally's HTML-based .xls files
                    if f.name.endswith('.xls'):
                        try:
                            df = pd.read_excel(f, engine='openpyxl')
                        except:
                            f.seek(0)
                            df = pd.read_html(f.read().decode('utf-8'))[0]
                    else:
                        df = pd.read_excel(f)
                        
                    current_item = None
                    
                    for idx, row in df.iterrows():
                        col0, col1 = row.iloc[0], row.iloc[1]
                        vch_type = str(row.iloc[2]).strip().lower() if pd.notna(row.iloc[2]) else ''
                        
                        # Identify item name
                        if pd.isna(col0) and pd.notna(col1):
                            val1 = str(col1).strip()
                            if val1 and val1 not in ['Particulars', 'Quantity']:
                                current_item = val1
                                if current_item not in metrics_dict:
                                    metrics_dict[current_item] = {
                                        'inward_production': 0.0,
                                        'inward_purchase': 0.0,
                                        'outward_consumption': 0.0,
                                        'outward_sales': 0.0,
                                        'godown_transfer': 0.0
                                    }
                                continue
                        
                        # Aggregate quantities by original clean rules
                        if current_item and vch_type:
                            in_qty = clean_val(row.iloc[4])
                            out_qty = clean_val(row.iloc[7])
                            
                            if 'conversion' in vch_type and in_qty > 0:
                                metrics_dict[current_item]['inward_production'] += in_qty
                            if 'purchase' in vch_type and in_qty > 0:
                                metrics_dict[current_item]['inward_purchase'] += in_qty
                            if 'conversion' in vch_type and out_qty > 0:
                                metrics_dict[current_item]['outward_consumption'] += out_qty
                            if 'sale' in vch_type and out_qty > 0:
                                metrics_dict[current_item]['outward_sales'] += out_qty
                            if ('warehouse' in vch_type or 'transfer' in vch_type) and in_qty > 0:
                                metrics_dict[current_item]['godown_transfer'] += in_qty
                                
                except Exception as e:
                    st.warning(f"Could not process {f.name}: {e}")

            # --- Step 2: Update Master Sheet ---
            master_file.seek(0)
            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            updates = 0
            
            # Start updating from row 12
            for row in range(12, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=1).value
                if cell_val and isinstance(cell_val, str):
                    item_name = cell_val.strip()
                    if item_name in metrics_dict:
                        m = metrics_dict[item_name]
                        
                        # Write calculated values if they are greater than 0
                        if m['inward_production'] > 0: ws.cell(row=row, column=18).value = m['inward_production']
                        if m['inward_purchase'] > 0: ws.cell(row=row, column=19).value = m['inward_purchase']
                        if m['outward_consumption'] > 0: ws.cell(row=row, column=20).value = m['outward_consumption']
                        if m['outward_sales'] > 0: ws.cell(row=row, column=21).value = m['outward_sales']
                        if m['godown_transfer'] > 0: ws.cell(row=row, column=24).value = m['godown_transfer']
                            
                        updates += 1

            # --- Step 3: Output for Download ---
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ Successfully processed files and updated {updates} items in the master sheet!")
            
            st.download_button(
                label="📥 Download Final Ledger",
                data=output,
                file_name="final_StkSum_Complete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
