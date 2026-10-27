import streamlit as st
import pandas as pd
import openpyxl
import math
import io

# --- Page Configuration ---
st.set_page_config(page_title="Tally Ledger Segregator", layout="centered", page_icon="📊")
st.title("📊 Tally Stock Ledger Segregator")
st.markdown("Upload your Master Sheet and your raw Tally exports to automatically calculate and update the ledger based on your custom rules.")

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
        with st.spinner("Processing files and applying rules..."):
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
                                        'sale_out': 0.0, 'credit_note_in': 0.0,
                                        'debit_note_in': 0.0, 'debit_note_out': 0.0,
                                        'wh_in': 0.0, 'wh_out': 0.0,
                                        'sj_conv_in': 0.0, 'sj_conv_out': 0.0,
                                        'sj_other_in': 0.0, 'sj_other_out': 0.0,
                                        'purchase_in': 0.0
                                    }
                                continue
                        
                        # Aggregate quantities by Voucher Type
                        if current_item and vch_type:
                            in_qty = clean_val(row.iloc[4])
                            out_qty = clean_val(row.iloc[7])
                            
                            # Sales and Credit Notes
                            if 'sale' in vch_type and out_qty > 0:
                                metrics_dict[current_item]['sale_out'] += out_qty
                            if 'credit note' in vch_type and in_qty > 0:
                                metrics_dict[current_item]['credit_note_in'] += in_qty
                                
                            # Debit Notes
                            if 'debit note' in vch_type:
                                if in_qty > 0: metrics_dict[current_item]['debit_note_in'] += in_qty
                                if out_qty > 0: metrics_dict[current_item]['debit_note_out'] += out_qty
                                    
                            # Warehouse Transfers
                            if 'warehouse' in vch_type or 'transfer' in vch_type:
                                if in_qty > 0: metrics_dict[current_item]['wh_in'] += in_qty
                                if out_qty > 0: metrics_dict[current_item]['wh_out'] += out_qty
                                    
                            # Stock Journal Conversion
                            elif 'conversion' in vch_type:
                                if in_qty > 0: metrics_dict[current_item]['sj_conv_in'] += in_qty
                                if out_qty > 0: metrics_dict[current_item]['sj_conv_out'] += out_qty
                                    
                            # Other Stock Journals
                            elif 'stock journal' in vch_type:
                                if in_qty > 0: metrics_dict[current_item]['sj_other_in'] += in_qty
                                if out_qty > 0: metrics_dict[current_item]['sj_other_out'] += out_qty
                                    
                            # Purchases
                            if 'purchase' in vch_type and 'debit note' not in vch_type and in_qty > 0:
                                metrics_dict[current_item]['purchase_in'] += in_qty
                                
                except Exception as e:
                    st.warning(f"Could not process {f.name}: {e}")

            # --- Step 2: Apply Custom Rules ---
            final_metrics = {}
            for item, d in metrics_dict.items():
                # Rule 1: Net Sale Qty (Outward Sales - Credit Notes)
                net_sale = d['sale_out'] - d['credit_note_in']
                
                # Rule 2: Warehouse trf in and out difference
                wh_in = d['wh_in']
                wh_out = d['wh_out']
                wh_diff = wh_in - wh_out
                godown_transfer = wh_in 
                
                # Rule 3: SJ net calculation
                total_in = d['sj_conv_in'] + d['sj_other_in'] + wh_diff + d['debit_note_in']
                total_out = d['sj_conv_out'] + d['sj_other_out'] + d['debit_note_out']
                
                net_prod = total_in - total_out
                
                inward_prod = net_prod if net_prod > 0 else 0.0
                outward_cons = abs(net_prod) if net_prod < 0 else 0.0
                
                final_metrics[item] = {
                    'net_sale': net_sale if net_sale > 0 else 0.0,
                    'godown_transfer': godown_transfer if godown_transfer > 0 else 0.0,
                    'inward_prod': inward_prod,
                    'outward_cons': outward_cons,
                    'inward_purchase': d['purchase_in']
                }

            # --- Step 3: Update Master Sheet ---
            master_file.seek(0)
            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            updates = 0
            
            # Start updating from row 12
            for row in range(12, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=1).value
                if cell_val and isinstance(cell_val, str):
                    item_name = cell_val.strip()
                    if item_name in final_metrics:
                        m = final_metrics[item_name]
                        
                        # Clear existing target columns to overwrite cleanly
                        for col in [18, 19, 20, 21, 24]: # R, S, T, U, X
                            ws.cell(row=row, column=col).value = None
                            
                        # Write calculated values
                        if m['inward_prod'] > 0: ws.cell(row=row, column=18).value = m['inward_prod']
                        if m['inward_purchase'] > 0: ws.cell(row=row, column=19).value = m['inward_purchase']
                        if m['outward_cons'] > 0: ws.cell(row=row, column=20).value = m['outward_cons']
                        if m['net_sale'] > 0: ws.cell(row=row, column=21).value = m['net_sale']
                        if m['godown_transfer'] > 0: ws.cell(row=row, column=24).value = m['godown_transfer']
                            
                        updates += 1

            # --- Step 4: Output for Download ---
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ Successfully processed {len(tally_files)} files and updated {updates} items in the master sheet!")
            
            st.download_button(
                label="📥 Download Segregated Ledger",
                data=output,
                file_name="Segregated_Summary_Ledger.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
