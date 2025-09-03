import streamlit as st
import pandas as pd
from io import StringIO

st.title("Credit Card Sales to QuickBooks IIF Converter")

uploaded_file = st.file_uploader("Upload Excel Report", type=["xlsx"])

if uploaded_file:
    # Load Excel
    df = pd.read_excel(uploaded_file, header=None)
    
    # Unmerge + drop header rows
    df = df.iloc[16:].reset_index(drop=True)
    
    # Select columns J (9), P (15), Z (25) - zero-based indexing
    df = df.iloc[:, [9, 15, 25]]
    df.columns = ["Date", "BillNo", "Amount"]
    
    # Clean fields
    df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%m/%d/%Y")
    df["BillNo"] = df["BillNo"].astype(str).str.strip()
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    
    # Create memo
    df["Memo"] = df["BillNo"].apply(lambda x: f"Mnarani Bill {x} Credit card sale")
    
    # Generate IIF
    iif = StringIO()
    iif.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\n")
    iif.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\n")
    iif.write("!ENDTRNS\n")
    
    for _, row in df.iterrows():
        iif.write(f"TRNS\tSALESREC\t{row['Date']}\tCredit Card Clearing\tWalk-In Customer\t{row['Amount']}\t{row['Memo']}\n")
        iif.write(f"SPL\tSALESREC\t{row['Date']}\tRevenue:POS Sales\tWalk-In Customer\t{-row['Amount']}\t{row['Memo']}\n")
        iif.write("ENDTRNS\n")
    
    # Download button
    st.download_button(
        label="Download IIF File",
        data=iif.getvalue(),
        file_name="credit_card_sales.iif",
        mime="text/plain"
    )
    
    st.success("IIF file generated successfully.")
