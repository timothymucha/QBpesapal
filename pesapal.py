import streamlit as st
import pandas as pd
from io import StringIO

st.title("Credit Card Sales to QuickBooks IIF Converter")

def truncate_at_blank(df):
    """Truncate at first completely blank row"""
    first_blank = df[df.isnull().all(axis=1)].index
    if not first_blank.empty:
        return df.loc[: first_blank[0] - 1]
    return df

uploaded_file = st.file_uploader("Upload Excel Report", type=["xlsx"])

if uploaded_file:
    try:
        # Read raw data starting from row 17 (skip first 16 rows)
        df_raw = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Rename relevant columns
        df_raw.rename(columns={
            4: "Till No",
            9: "Date",
            15: "Bill No.",
            25: "Amount"
        }, inplace=True)

        # Keep only relevant cols
        df = df_raw[["Till No", "Date", "Bill No.", "Amount"]].copy()

        # Truncate at first blank row
        df = truncate_at_blank(df)

        # Filter only Till rows (credit card tills e.g., MT01)
        df = df[df["Till No"].astype(str).str.contains("MT01", na=False)]

        # Clean Date
        df["Date"] = pd.to_datetime(
            df["Date"], format="%d-%b-%Y %I.%M.%S %p", errors="coerce"
        )
        df = df.dropna(subset=["Date"])
        df["Date"] = df["Date"].dt.strftime("%m/%d/%Y")

        # Clean Amount
        df["Amount"] = (
            df["Amount"].astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
        df = df.dropna(subset=["Amount"])

        # Create Memo
        df["Memo"] = df.apply(
            lambda x: f"Till {x['Till No']} | Invoice {x['Bill No.']}",
            axis=1
        )

        # --- Preview ---
        st.subheader("🧾 Preview: First 10 Cleaned Credit Card Sales")
        st.dataframe(df.head(10))

        # --- Generate IIF ---
        iif = StringIO()
        iif.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\n")
        iif.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\n")
        iif.write("!ENDTRNS\n")

        for _, row in df.iterrows():
            iif.write(
                f"TRNS\tPAYMENT\t{row['Date']}\tCredit Card Clearing\tWalk-In Customer\t{row['Amount']}\t{row['Memo']}\t{row['Bill No.']}\n"
            )
            iif.write(
                f"SPL\tPAYMENT\t{row['Date']}\tAccounts Receivable\tWalk-In Customer\t{-row['Amount']}\t{row['Memo']}\t\n"
            )
            iif.write("ENDTRNS\n")

        # Download
        st.download_button(
            label="⬇️ Download IIF File",
            data=iif.getvalue(),
            file_name="credit_card_sales.iif",
            mime="text/plain"
        )

        st.success("✅ IIF file generated successfully.")

    except Exception as e:
        st.error(f"Error processing file: {e}")
