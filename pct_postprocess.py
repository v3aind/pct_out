import streamlit as st
import pandas as pd

# Set Streamlit page title
st.title("Excel File Processor")

# Step 1: Get User Inputs
st.sidebar.header("User Inputs")

POID = st.sidebar.text_input("Enter the POID:")
ID = st.sidebar.text_input("Enter the ID:")
po_name = st.sidebar.text_input("Enter PO Name:")
master_keyword = st.sidebar.text_input("Enter Master Keyword:")

# File Upload: Upload the input Excel file
uploaded_file = st.file_uploader(
    "Upload the input Excel file",
    type=["xlsx"],
    help="Upload the input file (e.g., RESULT- ALL RULES -Prodef.xlsx)",
)

# Step 2: Generate Output File
if st.button("Generate Excel File"):
    if uploaded_file and POID and ID and po_name and master_keyword:
        # Load the input file
        file1 = pd.ExcelFile(uploaded_file)

        # Create the output file name
        output_file_name = f"PLD_{ID}_{POID}.xlsx"

        # Create an ExcelWriter object
        with pd.ExcelWriter(output_file_name, engine="xlsxwriter") as writer:
            # Create the "PO" sheet
            po_df = pd.DataFrame(
                {
                    "PO ID": [POID],
                    "PO Name": [po_name],
                    "Master Keyword": [master_keyword],
                    "Family": ["roamingSingleCountry"],
                    "PO Type": ["ADDON"],
                    "Product Category": ["b2cMobile"],
                    "Payment Type": ["Prepaid,Postpaid"],
                    "Action": ["NO_CHANGE"],
                }
            )
            po_df.to_excel(writer, sheet_name="PO", index=False)

            # Process other sheets
            sheet_names = [
                "Rules-Keyword",
                "Rules-Alias",
                "Rules-Header",
                "Rules-PCRF",
            ]

            for sheet_name in sheet_names:
                df = pd.read_excel(file1, sheet_name=sheet_name)
                if sheet_name == "Rules-PCRF":
                    df.to_excel(writer, sheet_name="PCRF", index=False)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Handle specific sheets
            try:
                df = pd.read_excel(file1, sheet_name="Rules-Cases-Condition")
                if "OpIndex" in df.columns:
                    df["OpIndex"] = pd.to_numeric(df["OpIndex"], errors="coerce").astype("Int64")
                df.to_excel(writer, sheet_name="Rules-Cases-Condition", index=False)
            except Exception as e:
                st.error(f"Error processing 'Rules-Cases-Condition': {e}")

            # Rules-Cases-Success
            try:
                df = pd.read_excel(file1, sheet_name="Rules-Cases-Success")
                if "OpIndex" in df.columns:
                    df["OpIndex"] = pd.to_numeric(df["OpIndex"], errors="coerce").astype("Int64")
                if "Ruleset ShortName" in df.columns:
                    df["Exit Value"] = df["Ruleset ShortName"].apply(
                        lambda x: "1" if pd.notna(x) and x.strip() != "" else ""
                    )
                df.to_excel(writer, sheet_name="Rules-Cases-Success", index=False)
            except Exception as e:
                st.error(f"Error processing 'Rules-Cases-Success': {e}")

            # Add placeholders for other sheets with "sample" data
            sample_sheets = {
                "Rules-GSI GRP Pack": ["Ruleset ShortName", "GSI GRP Pack-Group ID", "Action"],
                "Blacklist-Gift-Promocodes": ["Ruleset ShortName", "Coherence Key", "Promo Codes", "Action"],
                "MYIM3-UNREG": ["Ruleset ShortName", "Keyword", "Shortcode", "Unreg Flag", "Buy Extra Flag", "Action"],
            }

            for sheet_name, columns in sample_sheets.items():
                sample_df = pd.DataFrame([{col: "sample" for col in columns}])
                sample_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Step 3: Provide Download Option
        with open(output_file_name, "rb") as file:
            st.success(f"File '{output_file_name}' generated successfully!")
            st.download_button(
                label="Download Excel File",
                data=file,
                file_name=output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.error("Please fill in all required inputs and upload a valid file.")
