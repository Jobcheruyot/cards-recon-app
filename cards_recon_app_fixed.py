
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

        # Fix: Clean aspire selection block indentation
        aspire = aspire[[
            'STORE_CODE',
            'STORE_NAME',
            'ZED_DATE',
            'TILL',
            'SESSION',
            'RCT',
            'CUSTOMER_NAME',
            'CARD_TYPE',
            'CARD_NUMBER',
            'AMOUNT',
            'REF_NO',
            'RCT_TRN_DATE'
        ]]

        # Demo logic placeholder
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            kcb.to_excel(writer, index=False, sheet_name="KCB")
            equity.to_excel(writer, index=False, sheet_name="Equity")
            aspire.to_excel(writer, index=False, sheet_name="Aspire")
            key.to_excel(writer, index=False, sheet_name="Key")

        st.success("✅ Processing Complete!")
        st.download_button(
            label="📥 Download Reconciliation Report",
            data=output.getvalue(),
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
