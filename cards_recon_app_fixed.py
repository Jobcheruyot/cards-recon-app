# Prepare the corrected Streamlit app with full logic integration
streamlit_app_code = """
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
        # Step 1: Load files
        kcb = pd.read_excel(kcb_file)
        equity = pd.read_excel(equity_file)
        aspire = pd.read_csv(aspire_file)
        card_key = pd.read_excel(key_file)

        # Step 2: Clean card data
        for df in [kcb, equity]:
            df['Card_Number'] = df['Card_Number'].astype(str)
            df['card_check'] = df['Card_Number'].str[:4] + df['Card_Number'].str[-4:]
            df['Amount_check'] = df['Purchase'].round(2) == df['Amount'].round(2)

        merged_cards = pd.concat([kcb, equity], ignore_index=True)
        merged_cards['Source'] = merged_cards['Source'].fillna('UNKNOWN')
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str)
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        merged_cards['Amount_check'] = merged_cards['Purchase'].round(2) == merged_cards['Amount'].round(2)

        # Step 3: Clean aspire data
        aspire['card_check'] = aspire['CARD_NUMBER'].astype(str).str[:4] + aspire['CARD_NUMBER'].astype(str).str[-4:]
        aspire['REF_NO'] = aspire['REF_NO'].astype(str).str.lstrip('0')
        aspire['rrn_check'] = aspire['REF_NO'].map(dict(zip(merged_cards['R_R_N'].astype(str).str.lstrip('0'), merged_cards['Purchase'])))
        aspire['val_check'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce') - pd.to_numeric(aspire['rrn_check'], errors='coerce')

        # Step 4: Merge key to aspire
        aspire = aspire.merge(card_key[['store', 'Col_2']], how='left', left_on='STORE_NAME', right_on='store')
        aspire['branch'] = aspire['Col_2']
        aspire['Purchase'] = aspire['Purchase'].astype(str).str.split('.').str[0]
        aspire['Check_Two'] = aspire['branch'] + aspire['Purchase']

        # Step 5: Reconciliation flags
        merged_cards['card_check'] = merged_cards['card_check'].astype(str)
        aspire['card_check'] = aspire['card_check'].astype(str)
        aspire['card_conf'] = aspire.apply(
            lambda row: merged_cards.loc[
                (merged_cards['card_check'] == row['card_check']) & (row['rrn_check'] == 0),
                'Purchase'
            ].values[0] if len(
                merged_cards.loc[
                    (merged_cards['card_check'] == row['card_check']) & (row['rrn_check'] == 0)
                ]
            ) > 0 else '', axis=1)

        # Step 6: Filter and organize outputs
        aspire = aspire.drop_duplicates()
        unmatched_kcb = kcb[kcb['Amount_check'] == False]
        unmatched_equity = equity[equity['Amount_check'] == False]

        # Output Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aspire.to_excel(writer, index=False, sheet_name="Aspire_Cleaned")
            merged_cards.to_excel(writer, index=False, sheet_name="Merged_Cards")
            unmatched_kcb.to_excel(writer, index=False, sheet_name="Unmatched_KCB")
            unmatched_equity.to_excel(writer, index=False, sheet_name="Unmatched_Equity")
            card_key.to_excel(writer, index=False, sheet_name="Card_Key")

        st.success("✅ Reconciliation Complete!")

        st.download_button(
            label="📥 Download Final Reconciliation Report",
            data=output.getvalue(),
            file_name="Cards_Reconciliation_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
"""

# Save the complete working app to file
final_app_path = Path("/mnt/data/cards_recon_app_final.py")
final_app_path.write_text(streamlit_app_code)

final_app_path.name
