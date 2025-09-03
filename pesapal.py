import streamlit as st
import pandas as pd
from io import StringIO

st.title("Credit Card Sales to QuickBooks IIF Converter")

uploaded_file = st.file_uploader("Upload Excel Report", type=["xlsx"])

if uploaded_file:
    # Read Excel using row 13 (index 12) as header
    df = pd.read_excel(uploaded_file, header=12)
    
    # Keep only rows where Billing Station (column 'Till#' or col E) contains "MT01"
    if "Till#" in df.columns:
        df = df[df["Till#"].astype(str).str.contains("MT01", na=False)]
    else:
        st.error("Could not find 'Till#' column (expected in column E). Please check your file.")
        st.stop()
    
    # Keep only relevant columns
    df = df[["Bill Date", "Bill No.", "Amount"]].copy()
    
    # --- Clean fields ---
    # Parse and format Bill Date
    df["Bill Date"] = pd.to_datetime(df["Bill Date"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["Bill Date"])
    df["Date"] = df["Bill Date"].dt.strftime("%m/%d/%Y")
    
    # Clean Bill Number
    df["BillNo"] = df["Bill No."].astype(str).str.strip()
    
    # Clean Amount (remove commas / KES text, then numeric)
    df["Amount"] = (
        df["Amount"].astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("KES", "", regex=False)
        .str.strip()
    )
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df.dropna(subset=["Amount"])
    
    # Create Memo
    df["Memo"] = df["BillNo"].apply(lambda x: f"Mnarani Bill {x} Credit card sale")
    
    # Show preview
    st.subheader("Preview of Cleaned Data")
    st.dataframe(df[["Date", "BillNo", "Amount", "Memo"]].head(20))
    
    # Generate IIF content
    iif = StringIO()
    iif.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\n")
    iif.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\n")
    iif.write("!ENDTRNS\n")
    
    for _, row in df.iterrows():
        iif.write(
            f"TRNS\tSALESREC\t{row['Date']}\tCredit Card Clearing\tWalk-In Customer\t{row['Amount']}\t{row['Memo']}\n"
        )
        iif.write(
            f"SPL\tSALESREC\t{row['Date']}\tRevenue:POS Sales\tWalk-In Customer\t{-row['Amount']}\t{row['Memo']}\n"
        )
        iif.write("ENDTRNS\n")
    
    # Download button
    st.download_button(
        label="Download IIF File",
        data=iif.getvalue(),
        file_name="credit_card_sales.iif",
        mime="text/plain"
    )
    
    st.success("IIF file generated successfully.")
