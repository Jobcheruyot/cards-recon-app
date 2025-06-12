
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Cards Reconciliation Engine", layout="wide")
st.title("📊 Cards Reconciliation Report Generator")

# Upload input files
kcb_file = st.file_uploader("Upload KCB Excel", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel", type=["xlsx"])
aspire_file = st.file_uploader("Upload Aspire CSV", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel", type=["xlsx"])

if st.button("▶️ Generate Reconciliation Report"):
    if not all([kcb_file, equity_file, aspire_file, key_file]):
        st.warning("⚠️ Please upload all 4 input files.")
    else:
        # Step 1: Load files
        kcb = pd.read_excel(kcb_file)
        equity = pd.read_excel(equity_file)
        aspire = pd.read_csv(aspire_file)
        card_key = pd.read_excel(key_file)

        # Step 2: Clean KCB and Equity data
        for df in [kcb, equity]:
            df['Card_Number'] = df['Card_Number'].astype(str)
            df['card_check'] = df['Card_Number'].str[:4] + df['Card_Number'].str[-4:]
            df['Amount_check'] = df['Purchase'].round(2) == df['Amount'].round(2)

        # Step 3: Combine KCB and Equity
        merged_cards = pd.concat([kcb, equity], ignore_index=True)
        merged_cards['Source'] = merged_cards['Source'].fillna('UNKNOWN')
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        merged_cards['Amount_check'] = merged_cards['Purchase'].round(2) == merged_cards['Amount'].round(2)

        # Step 4: Prepare Aspire data
        aspire['REF_NO'] = aspire['REF_NO'].astype(str).str.lstrip("0")
        aspire['card_check'] = aspire['CARD_NUMBER'].astype(str).str[:4] + aspire['CARD_NUMBER'].astype(str).str[-4:]

        # Step 5: Match RRN and get Purchase
        merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.lstrip("0")
        merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')
        aspire['rrn_check'] = aspire['REF_NO'].map(dict(zip(merged_cards['R_R_N'], merged_cards['Purchase'])))
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        aspire['val_check'] = aspire['AMOUNT'] - pd.to_numeric(aspire['rrn_check'], errors='coerce')

        # Step 6: Enrich with card_key
        card_key.columns = ['store', 'branch']
        aspire = aspire.merge(card_key, how='left', left_on='STORE_NAME', right_on='store')
        aspire['Purchase'] = aspire['Purchase'].astype(str).str.split('.').str[0]
        aspire['Check_Two'] = aspire['branch'] + aspire['Purchase']

        # Step 7: Flag confirmation
        aspire['card_conf'] = aspire.apply(
            lambda row: merged_cards.loc[
                (merged_cards['card_check'] == row['card_check']) & (row['rrn_check'] == 0),
                'Purchase'
            ].values[0] if len(
                merged_cards.loc[
                    (merged_cards['card_check'] == row['card_check']) & (row['rrn_check'] == 0)
                ]
            ) > 0 else '', axis=1)

        # Output file generation
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            kcb.to_excel(writer, index=False, sheet_name='KCB')
            equity.to_excel(writer, index=False, sheet_name='Equity')
            aspire.to_excel(writer, index=False, sheet_name='Aspire')
            card_key.to_excel(writer, index=False, sheet_name='Key')
            merged_cards.to_excel(writer, index=False, sheet_name='Merged_Cards')

        st.success("✅ Reconciliation Report Generated")
        st.download_button(
            label="📥 Download Reconciliation Report",
            data=output.getvalue(),
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
