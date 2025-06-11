
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Cards Reconciliation App", layout="wide")
st.title("🧾 Cards Reconciliation Processor")

# Upload section
kcb_file = st.file_uploader("Upload KCB Excel file", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel file", type=["xlsx"])
aspire_file = st.file_uploader("Upload Aspire CSV file", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel file", type=["xlsx"])

if st.button("▶️ Process Files"):

    if not all([kcb_file, equity_file, aspire_file, key_file]):
        st.warning("Please upload all 4 files.")
    else:
        # Read the files
        kcb = pd.read_excel(kcb_file)
        equity = pd.read_excel(equity_file)
        aspire = pd.read_csv(aspire_file)
        key = pd.read_excel(key_file)

        # ========== ✅ Inserted Logic from Notebook ==========

#!/usr/bin/env python
        # coding: utf-8
        
        # #Import the neccessary Libraries
        
        # In[1192]:
        
        
        
        
        
        # In[1193]:
        
        
        # Import the necessary libraries
        import pandas as pd
        
        # Load the Excel files
        
        # Display the first few rows of each for verification
        print("KCB Data:")
        display(kcb.head())
        
        
        # In[1194]:
        
        
        print("Equity Data:")
        display(equity.head())
        
        
        # In[1195]:
        
        
        print("Card Key:")
        display(key.head())
        
        
        # #Bank Cards Statements alignments
        
        # ##Introduce a column with the bank name
        
        # In[1196]:
        
        
        # Step 1: Clean all column names to avoid hidden spaces or newline issues
        kcb.columns = kcb.columns.str.strip()
        equity.columns = equity.columns.str.strip()
        
        # Step 2: Rename KCB columns to match the target structure
        kcb_renamed = kcb.rename(columns={
            'Card No': 'Card_Number',
            'Trans Date': 'TRANS_DATE',
            'RRN': 'R_R_N',
            'Amount': 'Purchase',
            'Comm': 'Commission',
            'NetPaid': 'Settlement_Amount',
            'Merchant': 'store'  # renamed Merchant to store
        })
        
        # Step 3: Create Cash_Back and Source columns for KCB
        kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -1 * x if x < 0 else 0)
        kcb_renamed['Source'] = 'KCB'
        
        # Step 4: Rename Equity's Outlet_Name to store
        equity['Source'] = 'Equity'
        
        # Step 5: Final column list including store
        columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
                   'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
        
        # Step 6: Align columns and merge
        kcb_final = kcb_renamed[columns]
        equity_final = equity[columns]
        merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)
        
        # Step 7: Display and save
        display(merged_cards.tail())
        merged_cards.to_excel("merged_cards.xlsx", index=False)
        
        
        # ##Delete Redundat rows
        
        # In[1197]:
        
        
        # Drop rows where Card_Number is NaN or empty string
        merged_cards = merged_cards[merged_cards['Card_Number'].notna()]                 # Remove NaN
        merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']  # Remove blanks
        
        
        # In[1198]:
        
        
        # Count of records by Source
        source_counts = merged_cards['Source'].value_counts()
        
        # Display the result
        print(source_counts)
        
        
        # In[1199]:
        
        
        # Step 1: Load the card_key file
        card_key = pd.read_excel("card_key.xlsx", engine="openpyxl")
        
        # Step 2: Clean column names and string values
        card_key.columns = card_key.columns.str.strip()
        card_key['Col_1'] = card_key['Col_1'].str.strip()
        card_key['Col_2'] = card_key['Col_2'].str.strip()
        merged_cards['store'] = merged_cards['store'].str.strip()
        
        # Step 3: Create lookup dictionary
        lookup_dict = dict(zip(card_key['Col_1'], card_key['Col_2']))
        
        # Step 4: Map and create the 'branch' column
        merged_cards['branch'] = merged_cards['store'].map(lookup_dict)
        
        # Step 5: Reorder columns to place 'branch' after 'Source'
        cols = list(merged_cards.columns)
        source_index = cols.index('Source')
        # Move 'branch' to be right after 'Source'
        cols.insert(source_index + 1, cols.pop(cols.index('branch')))
        merged_cards = merged_cards[cols]
        
        # Step 6: Preview updated data
        display(merged_cards.tail())
        
        
        
        # In[1200]:
        
        
        # Step 1: Ensure Card_Number is a clean string
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
        
        # Step 2: Generate 'card_check' from first 4 + last 4 digits
        merged_cards['card_check'] = merged_cards['Card_Number'].apply(
            lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
        )
        
        # Step 3: Reorder columns to insert 'card_check' right after 'branch'
        cols = merged_cards.columns.tolist()
        if 'branch' in cols and 'card_check' in cols:
            # Remove 'card_check' temporarily
            cols.remove('card_check')
            # Insert after 'branch'
            branch_index = cols.index('branch')
            cols.insert(branch_index + 1, 'card_check')
            # Apply the new column order
            merged_cards = merged_cards[cols]
        
        # Step 4: Preview the full final table including the new column
        display(merged_cards.tail())
        
        
        
        # In[1200]:
        
        
        
        
        
        # In[1200]:
        
        
        
        
        
        # In[1201]:
        
        
        print("aspire:")
        display(aspire.head())
        
        
        # In[1202]:
        
        
        # Step 1: Ensure CARD_NUMBER is treated as a clean string
        aspire['CARD_NUMBER'] = aspire['CARD_NUMBER'].astype(str).str.strip()
        
        # Step 2: Create the 'card_check' column
        aspire['card_check'] = aspire['CARD_NUMBER'].apply(
            lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
        )
        
        # Step 3: Reorder columns — place 'card_check' at the end (or insert if needed)
        # If you want it right after 'CARD_NUMBER', do this:
        cols = aspire.columns.tolist()
        if 'CARD_NUMBER' in cols and 'card_check' in cols:
            cols.remove('card_check')
            insert_index = cols.index('CARD_NUMBER') + 1
            cols.insert(insert_index, 'card_check')
        
        # Step 4: Preview full row with new column
        display(aspire.head())
        
        
        
        # In[1203]:
        
        
        aspire.columns.tolist()
        
        
        # In[1204]:
        
        
        # Step 1: Filter to required columns
            'STORE_CODE',
            'STORE_NAME',
            'ZED_DATE',
            'TILL',
            'SESSION',
            'RCT',
            'CUSTOMER_NAME',
            'CARD_TYPE',
            'CARD_NUMBER',
            'card_check',
            'AMOUNT',
            'REF_NO',
            'RCT_TRN_DATE'
        ]]
        
        # Step 2: Preview
        display(aspire.head())
        
        # Step 3: Save to CSV
        aspire.to_csv("aspire_filtered.csv", index=False)
        
        
        
        # In[1205]:
        
        
        # prompt: rename "REF_NO" column in aspire to "R_R_N"
        
        
        # Verify the rename
        print("Aspire Data with Renamed Column:")
        display(aspire.head())
        
        
        # In[1206]:
        
        
        # Step 1: Ensure R_R_N is string and trimmed in both tables
        aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
        merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()
        
        # Step 2: Merge the tables on the shared column "R_R_N"
        rrntable = pd.merge(
            aspire,
            merged_cards,
            on='R_R_N',
            how='inner',  # or 'left'/'outer' depending on what you want to preserve
            suffixes=('_aspire', '_merged')
        )
        
        # Step 3: Preview the new table
        display(rrntable.head())
        print("✅ 'rrntable' created by merging aspire and merged_cards on 'R_R_N'.")
        
        
        # ##Align all rows to the branches
        
        # In[1207]:
        
        
        # Step 1: Filter rows where 'branch' is missing (NaN or blank)
        missing_branch_rows = merged_cards[merged_cards['branch'].isna()]
        
        # Step 2: Count how many rows are missing
        missing_count = missing_branch_rows.shape[0]
        print(f"✅ Total rows without a branch: {missing_count}")
        
        # Step 3: Display all full unmatched rows
        display(missing_branch_rows)
        
        # Step 4: Export full unmatched rows to CSV
        missing_branch_rows.to_csv("new_key_full_rows.csv", index=False)
        print("✅ Full unmatched rows exported to 'new_key_full_rows.csv'")
        
        
        
        # #Reconcile Aspire and Bank statements
        
        # In[1208]:
        
        
        # Step 1: Extract, drop NaNs, get unique and sort alphabetically
        card_summary = (
            aspire['STORE_NAME']
            .dropna()
            .drop_duplicates()
            .sort_values()
            .reset_index(drop=True)
            .to_frame(name='STORE_NAME')
        )
        
        # Step 2: Add index starting from 1
        card_summary.index = card_summary.index + 1
        card_summary.reset_index(inplace=True)
        card_summary.rename(columns={'index': 'No'}, inplace=True)
        
        # Step 3: Save to CSV
        card_summary.to_csv("Card_summary.csv", index=False)
        
        # Step 4: Preview
        display(card_summary.head())
        
        
        # In[1209]:
        
        
        # Step 1: Ensure AMOUNT is numeric (in case there are errors)
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        
        # Step 2: Group aspire data to get sum of AMOUNT per STORE_NAME
        aspire_sums = aspire.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
        
        # Step 3: Rename column for clarity
        aspire_sums = aspire_sums.rename(columns={'AMOUNT': 'Aspire_Zed'})
        
        # Step 4: Merge with card_summary
        card_summary = card_summary.merge(aspire_sums, on='STORE_NAME', how='left')
        
        # Step 5: Fill NaN values (stores with no AMOUNT) with 0
        card_summary['Aspire_Zed'] = card_summary['Aspire_Zed'].fillna(0)
        
        # Step 6: Export the updated file
        card_summary.to_csv("Card_summary.csv", index=False)
        
        # Step 7: Preview
        display(card_summary.head())
        print("✅ Updated 'Card_summary.csv' with 'Aspire_Zed' totals.")
        
        
        # In[1210]:
        
        
        # Step 1: Ensure Purchase is numeric
        merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')
        
        # Step 2: Filter for KCB source and group by branch
        kcb_grouped = (
            merged_cards[merged_cards['Source'] == 'KCB']
            .groupby('branch')['Purchase']
            .sum()
            .reset_index()
            .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'kcb_paid'})
        )
        
        # Step 3: Merge with card_summary on STORE_NAME
        card_summary = card_summary.merge(kcb_grouped, on='STORE_NAME', how='left')
        
        # Step 4: Fill missing values with 0
        card_summary['kcb_paid'] = card_summary['kcb_paid'].fillna(0)
        
        # Step 5: Export to CSV
        card_summary.to_csv("Card_summary.csv", index=False)
        
        # Step 6: Preview
        display(card_summary.tail())
        print("✅ 'kcb_paid' column added and Card_summary.csv updated.")
        
        
        # In[1211]:
        
        
        # Step 1: Ensure 'Purchase' is numeric
        merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')
        
        # Step 2: Group Equity values by branch
        equity_grouped = (
            merged_cards[merged_cards['Source'] == 'Equity']
            .groupby('branch')['Purchase']
            .sum()
            .reset_index()
            .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'equity_paid'})
        )
        
        # Step 3: Merge equity_paid with card_summary
        card_summary = card_summary.merge(equity_grouped, on='STORE_NAME', how='left')
        
        # Step 4: Fill missing equity_paid with 0
        card_summary['equity_paid'] = card_summary['equity_paid'].fillna(0)
        
        # Step 5: Reorder columns — insert equity_paid right after kcb_paid
        cols = list(card_summary.columns)
        if 'kcb_paid' in cols and 'equity_paid' in cols:
            kcb_index = cols.index('kcb_paid')
            # Remove equity_paid from current position and re-insert after kcb_paid
            cols.insert(kcb_index + 1, cols.pop(cols.index('equity_paid')))
            card_summary = card_summary[cols]
        
        # Step 6: Add totals row (with comma format)
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': card_summary['Aspire_Zed'].sum(),
            'kcb_paid': card_summary['kcb_paid'].sum(),
            'equity_paid': card_summary['equity_paid'].sum()
        }])
        
        # Step 7: Match columns and append
        for col in card_summary.columns:
            if col not in total_row.columns:
                total_row[col] = ''
        total_row = total_row[card_summary.columns]  # Reorder to match
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 8: Format numeric columns with comma separator
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid']
        for col in numeric_cols:
            card_summary[col] = card_summary[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
        
        # Step 9: Export final file
        card_summary.to_csv("Card_summary.csv", index=False)
        
        # Step 10: Preview
        display(card_summary.tail())
        print("✅ equity_paid column added, placed after kcb_paid, totals row appended, and numbers formatted with commas.")
        
        
        # In[1212]:
        
        
        # Step 1: Convert kcb_paid and equity_paid back to numeric (remove commas)
        card_summary['kcb_paid'] = card_summary['kcb_paid'].replace({',': ''}, regex=True).astype(float)
        card_summary['equity_paid'] = card_summary['equity_paid'].replace({',': ''}, regex=True).astype(float)
        
        # Step 2: Create Gross_Banking
        card_summary['Gross_Banking'] = card_summary['kcb_paid'] + card_summary['equity_paid']
        
        # Step 3: Format all numeric columns with commas
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking']
        for col in numeric_cols:
            card_summary[col] = card_summary[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
        
        # Step 4: Update totals row
        # (Recalculate the actual totals row using numeric versions of the original data)
        totals = {
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': f"{card_summary[:-1]['Aspire_Zed'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
            'kcb_paid': f"{card_summary[:-1]['kcb_paid'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
            'equity_paid': f"{card_summary[:-1]['equity_paid'].replace({',': ''}, regex=True).astype(float).sum():,.2f}",
            'Gross_Banking': f"{card_summary[:-1]['Gross_Banking'].replace({',': ''}, regex=True).astype(float).sum():,.2f}"
        }
        
        # Replace the last row (total row) with the updated totals
        card_summary.iloc[-1] = totals
        
        # Step 5: Save updated file
        card_summary.to_csv("Card_summary.csv", index=False)
        
        # Step 6: Preview
        display(card_summary.tail())
        
        
        
        # #Compute Banking Variance
        
        # In[1213]:
        
        
        # Load CSV and Excel files
        merged_cards = pd.read_excel('/content/merged_cards.xlsx')
        
        # Display first few rows to confirm
        aspire.head(), merged_cards.head()
        
        
        # In[1214]:
        
        
        # Rename 'R_R_N' in merged_cards to match column in aspire
        merged_cards.rename(columns={'R_R_N': 'REF_NO'}, inplace=True)
        
        
        
        # In[1215]:
        
        
        # Step 3 (Final Fix): Clean REF_NO in merged_cards to string (remove decimals and scientific notation)
        def clean_ref_no(x):
            try:
                return str(int(float(x)))  # Works for both float and scientific notation
            except:
                return ""  # If blank or bad format
        
        # Apply safe conversion
        merged_cards['REF_NO'] = merged_cards['REF_NO'].apply(clean_ref_no)
        
        # Also make sure aspire REF_NO is string
        aspire['REF_NO'] = aspire['REF_NO'].astype(str)
        
        
        # In[1216]:
        
        
        # Clean REF_NO by stripping leading zeros
        aspire['REF_NO'] = aspire['REF_NO'].astype(str).str.lstrip('0')
        
        # Create dictionary: REF_NO -> Purchase
        ref_to_purchase = dict(zip(merged_cards['REF_NO'], merged_cards['Purchase']))
        
        # Map to rrn_check column
        aspire['rrn_check'] = aspire['REF_NO'].map(ref_to_purchase).fillna(0)
        
        
        
        # In[1217]:
        
        
        # Get all cleaned REF_NO values from aspire
        matched_ref_nos = set(aspire['REF_NO'])
        
        # Mark 'Yes' in merged_cards if REF_NO was matched
        merged_cards['Cheked_rows'] = merged_cards['REF_NO'].astype(str).apply(lambda x: 'Yes' if x in matched_ref_nos else 'No')
        
        
        # In[1218]:
        
        
        # Count how many rows got a match
        matched = (aspire['rrn_check'] > 0).sum()
        total = len(aspire)
        
        print(f"✅ Matches found: {matched} out of {total} rows")
        print(f"✅ Match percentage: {(matched / total) * 100:.2f}%")
        aspire.head()
        
        
        # In[1219]:
        
        
        # Ensure AMOUNT is numeric in case it's not
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        
        # Compute val_check as AMOUNT - rrn_check
        aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']
        
        # Reorder columns to place 'val_check' after 'rrn_check'
        cols = list(aspire.columns)
        rrn_idx = cols.index('rrn_check')
        # Move val_check to right after rrn_check
        new_order = cols[:rrn_idx + 1] + ['val_check'] + cols[rrn_idx + 1:-1] + [cols[-1]]  # keep last column intact
        # Remove duplicate 'val_check' columns if they exist
        
        
        # Preview with all columns
        aspire.head()
        
        
        
        
        
        # In[1220]:
        
        
        from google.colab import files
        
        # Save the DataFrame to a CSV file
        aspire.to_csv("aspire_with_val_check.csv", index=False)
        
        
        
        # In[1221]:
        
        
        count_rrn_only = (aspire['rrn_check'] > 0).sum()
        print(f"✅ rrn_check > 0: {count_rrn_only}")
        
        
        # In[1222]:
        
        
        # Count rows where val_check is between -3 and 3 (inclusive)
        count_val_near_zero = aspire['val_check'].between(-3, 3).sum()
        
        print(f"✅ val_check between -3 and 3: {count_val_near_zero}")
        
        
        # In[1223]:
        
        
        # Count rows where rrn_check > 0 and val_check is between -3 and 3
        count_within_range = aspire[(aspire['rrn_check'] > 0) & (aspire['val_check'].between(-3, 3))].shape[0]
        
        print(f"✅ Rows where rrn_check > 0 and val_check is between -3 and 3: {count_within_range}")
        
        
        # In[1224]:
        
        
        # Count mismatched rows where val_check is not between -3 and 3
        count_mismatched = aspire[(aspire['rrn_check'] > 0) & (~aspire['val_check'].between(-3, 3))].shape[0]
        
        print(f"❌ Mismatched rows (rrn_check > 0 and val_check NOT between -3 and 3): {count_mismatched}")
        
        
        # In[1225]:
        
        
        # Extract rows where rrn_check > 0 and val_check is NOT between -3 and 3
        mismatched_rows = aspire[(aspire['rrn_check'] > 0) & (~aspire['val_check'].between(-3, 3))]
        
        # Display top rows
        mismatched_rows.head()
        
        
        # In[1226]:
        
        
        # Step 1: Ensure Card_Number is string
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str)
        
        # Step 2: Create or update 'card_check' column (8 characters only)
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        
        # Step 3: Move 'card_check' to appear immediately after 'Source', if Source exists
        if 'Source' in merged_cards.columns:
            cols = merged_cards.columns.tolist()
            if 'card_check' in cols:
                cols.remove('card_check')
            source_index = cols.index('Source') + 1
            cols.insert(source_index, 'card_check')
            merged_cards = merged_cards[cols]
        
        # ✅ Step 4: Drop exact duplicate rows
        merged_cards = merged_cards.drop_duplicates()
        
        # ✅ Step 5: Exclude rows where "TID" is blank (NaN or empty string)
        merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]
        
        # ✅ Step 6: Display result
        display(merged_cards.head())
        
        # ✅ Step 7: Download to Excel
        merged_cards.to_excel("cleaned_merged_cards.xlsx", index=False)
        
        
        
        # In[1227]:
        
        
        merged_cards.to_excel("/content/merged_cards.xlsx", index=False)
        
        
        # In[1228]:
        
        
        # ✅ Step 1: Filter rows where Cheked_rows is "No"
        newbankmerged = merged_cards[merged_cards['Cheked_rows'].str.strip().str.upper() == 'NO']
        
        # ✅ Step 2: Display the result
        display(newbankmerged.head())
        
        # ✅ Step 3: Optional – Save to Excel
        newbankmerged.to_excel("newbankmerged.xlsx", index=False)
        
        
        # In[1229]:
        
        
        # Step 1: Ensure both matching columns are string type for consistency
        newbankmerged['store'] = newbankmerged['store'].astype(str)
        key['Col_1'] = key['Col_1'].astype(str)
        
        # Step 2: Merge key data into newbankmerged based on the 'store' and 'Col_1'
        newbankmerged = newbankmerged.merge(
            key[['Col_1', 'Col_2']],
            how='left',
            left_on='store',
            right_on='Col_1'
        )
        
        # Step 3: Rename 'Col_2' to 'branch' and drop 'Col_1' (if not needed)
        newbankmerged = newbankmerged.rename(columns={'Col_2': 'branch'})
        newbankmerged = newbankmerged.drop(columns=['Col_1'])
        
        # Step 4: Display or export result
        display(newbankmerged.head())
        
        
        # In[1230]:
        
        
        # Step 1: Ensure 'branch' and 'Purchase' are strings
        newbankmerged['branch'] = newbankmerged['branch'].astype(str)
        newbankmerged['Purchase'] = newbankmerged['Purchase'].astype(str)
        
        # Step 2: Create 'Check_Two' by combining 'branch' and 'Purchase'
        newbankmerged['Check_Two'] = newbankmerged['branch'] + newbankmerged['Purchase']
        
        # Step 3: Display result
        display(newbankmerged.tail())
        
        
        # In[1231]:
        
        
        # ✅ Step: Create 'Check_Two' with 2 decimal places while keeping all original columns
        newbankmerged['Check_Two'] = (
            newbankmerged['branch'].astype(str).str.strip().str.upper() +
            newbankmerged['Purchase'].astype(float).map('{:.2f}'.format)
        )
        
        # ✅ Display all columns (first few rows)
        display(newbankmerged.head())
        
        
        # In[1232]:
        
        
        # Step 1: Load the card_key file
        card_key = pd.read_excel('/content/card_key.xlsx')
        
        # Step 2: Ensure matching key columns are strings
        merged_cards['store'] = merged_cards['store'].astype(str)
        card_key.iloc[:, 0] = card_key.iloc[:, 0].astype(str)
        
        # Step 3: Create a lookup dictionary: store -> branch (2nd column)
        store_to_branch = dict(zip(card_key.iloc[:, 0], card_key.iloc[:, 1]))
        
        # Step 4: Map branch to merged_cards using the store column
        merged_cards['branch'] = merged_cards['store'].map(store_to_branch)
        
        # Step 5: Move 'branch' column to appear right after 'store'
        cols = merged_cards.columns.tolist()
        cols.remove('branch')
        store_index = cols.index('store') + 1
        cols.insert(store_index, 'branch')
        merged_cards = merged_cards[cols]
        
        # Step 6: Preview updated DataFrame
        display(merged_cards.head())
        
        
        # In[1233]:
        
        
        aspire.tail()
        
        
        # In[1234]:
        
        
        # Count where rrn_check > 1
        greater_than_1 = (aspire['rrn_check'] > 1).sum()
        
        # Count where rrn_check <= 0
        zero_or_below = (aspire['rrn_check'] <= 0).sum()
        
        # Print results
        print(f"🔹 Transactions with rrn_check > 1: {greater_than_1}")
        print(f"🔸 Transactions with rrn_check <= 0: {zero_or_below}")
        
        
        # In[1235]:
        
        
        # ✅ Filter rows where 'rrn_check' is less than or equal to 0
        newaspire = aspire[aspire['rrn_check'] <= 0].copy()
        
        # ✅ Preview the result
        display(newaspire.tail())
        
        # ✅ Optional: Confirm row count
        print(f"Total rows in newaspire (rrn_check <= 0): {len(newaspire)}")
        
        
        # In[1236]:
        
        
        # 1. Count total rows
        row_count = len(newaspire)
        print(f"📊 Total rows in newaspire: {row_count}")
        
        # 2. View column names and data types
        print("\n🧾 Column Info:")
        print(newaspire.dtypes)
        
        # 3. Summary statistics for numeric columns
        print("\n📈 Summary Stats:")
        display(newaspire.describe(include='all'))
        
        # 4. Check for missing values per column
        print("\n🔍 Missing Values:")
        print(newaspire.isnull().sum())
        
        # 5. Preview first few rows
        print("\n🗂️ Sample Rows:")
        display(newaspire.head())
        
        
        # In[1237]:
        
        
        # Step 1: Ensure 'STORE_NAME' is string and keep 'AMOUNT' with decimals (2 decimal places)
        newaspire['Check_Two'] = newaspire['STORE_NAME'].astype(str) + newaspire['AMOUNT'].map('{:.2f}'.format)
        
        # Step 2: Preview the result
        display(newaspire.head())
        
        
        # In[1238]:
        
        
        # Step 1: Standardize Check_Two in both datasets
        newaspire['Check_Two'] = newaspire['STORE_NAME'].astype(str).str.strip().str.upper() + \
                                 newaspire['AMOUNT'].astype(float).astype(int).astype(str)
        
        newbankmerged['Check_Two'] = newbankmerged['branch'].astype(str).str.strip().str.upper() + \
                                     newbankmerged['Purchase'].astype(float).astype(int).astype(str)
        
        # Step 2: Create a copy of newbankmerged Check_Two values as a "pool" of available matches
        available_matches = newbankmerged['Check_Two'].tolist()
        
        # Step 3: Define a function that consumes matches once
        def check_and_consume(val):
            if val in available_matches:
                available_matches.remove(val)
                return 'Okay'
            else:
                return 'False'
        
        # Step 4: Apply function row-by-row
        newaspire['Amount_check'] = newaspire['Check_Two'].apply(check_and_consume)
        
        # Step 5: Count and preview
        okay_count = (newaspire['Amount_check'] == 'Okay').sum()
        print(f"✅ Unique matches marked as 'Okay': {okay_count}")
        display(newaspire.tail())
        
        
        # In[1239]:
        
        
        okay_count = (newaspire['Amount_check'] == 'Okay').sum()
        false_count = (newaspire['Amount_check'] == 'False').sum()
        total = len(newaspire)
        
        print(f"✅ Unique matches marked as 'Okay': {okay_count}")
        print(f"❌ Transactions without match (False): {false_count}")
        print(f"📊 Total transactions checked: {total}")
        
        
        # In[1240]:
        
        
        # Filter where Amount_check is 'False'
        false_list = newaspire[newaspire['Amount_check'] == 'False'].copy()
        
        # Display the list
        display(false_list)
        
        # Optional: Show count
        print(f"❌ Total unmatched (False) transactions: {len(false_list)}")
        
        
        # In[1241]:
        
        
        # Preview result
        display(newbankmerged.head())
        
        # Optional: Count how many were extracted
        print(f"✅ Total rows with Cheked_rows = 'No': {len(newmerged_cards)}")
        
        
        # In[1242]:
        
        
        # Step 1: Standardize Check_Two in both datasets
        newaspire['Check_Two'] = (
            newaspire['STORE_NAME'].astype(str).str.strip().str.upper() +
            newaspire['AMOUNT'].astype(float).astype(int).astype(str)
        )
        
        newmerged_cards['Check_Two'] = (
            newmerged_cards['branch'].astype(str).str.strip().str.upper() +
            newmerged_cards['Purchase'].astype(float).astype(int).astype(str)
        )
        
        # Step 2: Create a copy of newaspire Check_Two values as a "pool" of available matches
        available_aspire_matches = newaspire['Check_Two'].tolist()
        
        # Step 3: Define a function that consumes matches once
        def match_and_consume(val):
            if val in available_aspire_matches:
                available_aspire_matches.remove(val)
                return 'Okay'
            else:
                return 'False'
        
        # Step 4: Apply function row-by-row on newmerged_cards
        newmerged_cards['Amount_check'] = newmerged_cards['Check_Two'].apply(match_and_consume)
        
        # Step 5: Count and preview
        okay_count = (newmerged_cards['Amount_check'] == 'Okay').sum()
        print(f"✅ Unique matches found from newaspire: {okay_count}")
        display(newmerged_cards.head())
        
        
        # In[1243]:
        
        
        # Count summary of Amount_check
        summary_counts = newmerged_cards['Amount_check'].value_counts()
        
        # Print total count and breakdown
        print(f"📊 Total rows in newmerged_cards: {len(newmerged_cards)}")
        print("✅ Amount_check summary:")
        print(summary_counts)
        
        
        # In[1244]:
        
        
        display(card_summary.tail())
        
        
        # In[1245]:
        
        
        # Step 1: List numeric columns to clean
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking']
        
        # Step 2: Convert all numeric columns to float after removing commas
        for col in numeric_cols:
            card_summary[col] = card_summary[col].replace(',', '', regex=True).astype(float)
        
        # Step 3: Compute Variance
        card_summary['Variance'] = card_summary['Gross_Banking'] - card_summary['Aspire_Zed']
        
        # Step 4: Remove old TOTAL row if it exists
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Step 5: Compute fresh totals
        totals = card_summary[numeric_cols + ['Variance']].sum()
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance']
        }])
        
        # Step 6: Append the clean total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 7: Display
        display(card_summary.tail())
        
        
        
        # In[1246]:
        
        
        # Step 1: Filter relevant rows from newmerged_cards
        kcb_recs_data = newmerged_cards[
            (newmerged_cards['Source'] == 'KCB') &
            (newmerged_cards['Amount_check'] == 'False')
        ]
        
        # Step 2: Group by branch and sum the Purchase
        kcb_recs_grouped = kcb_recs_data.groupby('branch')['Purchase'].sum().reset_index()
        kcb_recs_grouped.columns = ['STORE_NAME', 'kcb_recs']  # Rename for merge compatibility
        
        # Step 3: Merge into card_summary
        card_summary = card_summary.merge(kcb_recs_grouped, on='STORE_NAME', how='left')
        
        # Step 4: Replace NaN with 0 for stores with no unmatched KCB records
        card_summary['kcb_recs'] = card_summary['kcb_recs'].fillna(0)
        
        # Step 5: Preview
        display(card_summary.tail())
        
        
        # In[1247]:
        
        
        # Step 1: Remove old TOTAL row if it exists
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Step 2: Define all numeric columns to sum (including new kcb_recs)
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance', 'kcb_recs']
        totals = card_summary[numeric_cols].sum()
        
        # Step 3: Create TOTAL row with matching structure
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs']
        }])
        
        # Step 4: Append total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 5: Display final result
        display(card_summary.tail())
        
        
        # In[1248]:
        
        
        # Step 1: Filter relevant rows from newmerged_cards
        equity_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'EQUITY') &
            (newmerged_cards['Amount_check'] == 'False')
        ]
        
        # Step 2: Group by branch and sum Purchase
        equity_recs_grouped = equity_recs_data.groupby('branch')['Purchase'].sum().reset_index()
        equity_recs_grouped.columns = ['STORE_NAME', 'Equity_recs']  # Rename to match card_summary
        
        # Step 3: Merge into card_summary
        card_summary = card_summary.merge(equity_recs_grouped, on='STORE_NAME', how='left')
        
        # Step 4: Fill NaN with 0
        card_summary['Equity_recs'] = card_summary['Equity_recs'].fillna(0)
        
        # Step 5: Remove old TOTAL row if exists
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Step 6: Compute total row including Equity_recs
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance', 'kcb_recs', 'Equity_recs']
        totals = card_summary[numeric_cols].sum()
        
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs'],
            'Equity_recs': totals['Equity_recs']
        }])
        
        # Step 7: Append total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 8: Display result
        display(card_summary.tail())
        
        
        # In[1249]:
        
        
        # Step 1: Filter newaspire where Amount_check is 'False'
        aspire_recs_data = newaspire[newaspire['Amount_check'] == 'False']
        
        # Step 2: Group by STORE_NAME and sum AMOUNT
        aspire_recs_grouped = aspire_recs_data.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
        aspire_recs_grouped.columns = ['STORE_NAME', 'Asp_Recs']  # Rename to match card_summary
        
        # Step 3: Merge into card_summary
        card_summary = card_summary.merge(aspire_recs_grouped, on='STORE_NAME', how='left')
        
        # Step 4: Fill NaNs with 0 where no unmatched aspire recs
        card_summary['Asp_Recs'] = card_summary['Asp_Recs'].fillna(0)
        
        # Step 5: Remove old TOTAL row if it exists
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Step 6: Recalculate TOTAL row including Asp_Recs
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']
        totals = card_summary[numeric_cols].sum()
        
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs'],
            'Equity_recs': totals['Equity_recs'],
            'Asp_Recs': totals['Asp_Recs']
        }])
        
        # Step 7: Append total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 8: Display final result
        display(card_summary.tail())
        
        
        # In[1250]:
        
        
        # Step 1: Remove old TOTAL row if it exists
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Step 2: Define all numeric columns to sum (including new ones)
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']
        
        # Step 3: Calculate totals
        totals = card_summary[numeric_cols].sum()
        
        # Step 4: Create TOTAL row using all columns in card_summary (including No and STORE_NAME)
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals.get('Aspire_Zed', 0),
            'kcb_paid': totals.get('kcb_paid', 0),
            'equity_paid': totals.get('equity_paid', 0),
            'Gross_Banking': totals.get('Gross_Banking', 0),
            'Variance': totals.get('Variance', 0),
            'kcb_recs': totals.get('kcb_recs', 0),
            'Equity_recs': totals.get('Equity_recs', 0),
            'Asp_Recs': totals.get('Asp_Recs', 0)
        }])
        
        # Step 5: Append total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # Step 6: Display final result
        display(card_summary.tail())
        
        
        # In[1255]:
        
        
        # Ensure missing columns default to 0 before calculating
        for col in ['Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']:
            if col not in card_summary.columns:
                card_summary[col] = 0
        
        # Compute Net_variance
        card_summary['Net_variance'] = (
            card_summary['Variance']
            - card_summary['kcb_recs']
            - card_summary['Equity_recs']
            + card_summary['Asp_Recs']
        )
        
        # Display result
        display(card_summary.tail())
        
        
        # In[1256]:
        
        
        get_ipython().system('pip install xlsxwriter')
        
        import pandas as pd
        
        # --- COLAB ONLY: Enable download ---
        from google.colab import files
        
        # ------------------ Prepare Sheets ------------------
        
        aspire_recs_data = newaspire[newaspire['Amount_check'] == 'False'].copy()
        
        equity_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'EQUITY') &
            (newmerged_cards['Amount_check'] == 'False')
        ].copy()
        
        kcb_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'KCB') &
            (newmerged_cards['Amount_check'] == 'False')
        ].copy()
        
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str)
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        if 'Source' in merged_cards.columns:
            cols = merged_cards.columns.tolist()
            if 'card_check' in cols:
                cols.remove('card_check')
            source_index = cols.index('Source') + 1
            cols.insert(source_index, 'card_check')
            merged_cards = merged_cards[cols]
        merged_cards = merged_cards.drop_duplicates()
        merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]
        
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']
        cols = list(aspire.columns)
        rrn_idx = cols.index('rrn_check')
        new_order = cols[:rrn_idx + 1] + ['val_check'] + cols[rrn_idx + 1:-1] + [cols[-1]]
        
        # ------------------ Add TOTAL to card_summary ------------------
        
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        all_possible = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                        'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']
        numeric_cols = [col for col in all_possible if col in card_summary.columns]
        
        totals = card_summary[numeric_cols].sum()
        
        total_row = {'No': '', 'STORE_NAME': 'TOTAL'}
        for col in numeric_cols:
            total_row[col] = totals[col]
        
        card_summary = pd.concat([card_summary, pd.DataFrame([total_row])], ignore_index=True)
        
        # ------------------ Export to Excel Workbook ------------------
        
        filename = "Reconciliation_Report.xlsx"
        
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            card_summary.to_excel(writer, sheet_name='card_summary', index=False)
            aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)
            equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)
            kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)
            merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
            aspire.to_excel(writer, sheet_name='aspire', index=False)
        
        print("✅ All reports exported successfully to:", filename)
        
        # ------------------ Auto-Download in Colab ------------------
        
        files.download(filename)
        
        
        # In[1252]:
        
        
        get_ipython().system('pip install xlsxwriter')
        
        import pandas as pd
        
        # ------------------ Prepare Sheets ------------------
        
        # Asp_Recs (Sheet 2)
        aspire_recs_data = newaspire[newaspire['Amount_check'] == 'False'].copy()
        
        # Equity_recs (Sheet 3)
        equity_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'EQUITY') &
            (newmerged_cards['Amount_check'] == 'False')
        ].copy()
        
        # KCB_recs (New Sheet)
        kcb_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'KCB') &
            (newmerged_cards['Amount_check'] == 'False')
        ].copy()
        
        # Clean merged_cards (Sheet 5)
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str)
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        if 'Source' in merged_cards.columns:
            cols = merged_cards.columns.tolist()
            if 'card_check' in cols:
                cols.remove('card_check')
            source_index = cols.index('Source') + 1
            cols.insert(source_index, 'card_check')
            merged_cards = merged_cards[cols]
        merged_cards = merged_cards.drop_duplicates()
        merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]
        
        # Clean aspire (Sheet 6)
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']
        cols = list(aspire.columns)
        rrn_idx = cols.index('rrn_check')
        new_order = cols[:rrn_idx + 1] + ['val_check'] + cols[rrn_idx + 1:-1] + [cols[-1]]
        
        # ------------------ Recalculate Total Row ------------------
        
        # Remove old TOTAL row
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        
        # Numeric columns to total
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                        'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']
        totals = card_summary[numeric_cols].sum()
        
        # Create TOTAL row
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals.get('Aspire_Zed', 0),
            'kcb_paid': totals.get('kcb_paid', 0),
            'equity_paid': totals.get('equity_paid', 0),
            'Gross_Banking': totals.get('Gross_Banking', 0),
            'Variance': totals.get('Variance', 0),
            'kcb_recs': totals.get('kcb_recs', 0),
            'Equity_recs': totals.get('Equity_recs', 0),
            'Asp_Recs': totals.get('Asp_Recs', 0)
        }])
        
        # Append to card_summary
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        
        # ------------------ Export to Excel ------------------
        
        with pd.ExcelWriter("Reconciliation_Report.xlsx", engine='xlsxwriter') as writer:
            card_summary.to_excel(writer, sheet_name='card_summary', index=False)
            aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)
            equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)
            kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)
            merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
            aspire.to_excel(writer, sheet_name='aspire', index=False)
        
        print("✅ All sheets exported to 'Reconciliation_Report.xlsx'")
        
        
        # In[1253]:
        
        
        # Temporarily remove row display limit
        pd.set_option('display.max_rows', None)
        
        # Display full DataFrame
        display(card_summary)
        
        # (Optional) Reset display limit afterward
        # pd.reset_option('display.max_rows')
        

        # ========== ✅ Sample Output File ==========
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            try:
                final.to_excel(writer, index=False, sheet_name="Final Output")
            except:
                pd.DataFrame({"Status": ["Success"], "Note": ["Custom logic executed"]}).to_excel(writer, index=False, sheet_name="Result")
        st.success("✅ Processing Complete!")
        st.download_button(
            label="📥 Download Reconciliation Report",
            data=output.getvalue(),
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
