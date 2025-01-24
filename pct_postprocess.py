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

            # Example sheet creation: Rules-Messages
            messages_df = pd.DataFrame(
                {
                    "PO ID": ["sample"],
                    "Ruleset ShortName": ["sample"],
                    "Order Status": ["sample"],
                    "Order Type": ["sample"],
                    "Sender Address": ["sample"],
                    "Channel": ["sample"],
                    "Message Content Index": ["sample"],
                    "Message Content": ["sample"],
                    "Action": ["NO CHANGE"],
                }
            )
            messages_df.to_excel(writer, sheet_name="Rules-Messages", index=False)

            # Sheet 9: Rules-Price-Mapping
            df = pd.read_excel(file1, sheet_name="Rules-Price-Mapping")

            # Convert "Price" column to integers
            df["Price"] = pd.to_numeric(df["Price"], errors="coerce").astype("Int64")

            # Ensure the "SID" column exists and manipulate it as needed
            if "SID" in df.columns:
                # Convert to string and strip whitespace, replace NaN with empty strings
                df["SID"] = df["SID"].astype(str).str.strip().replace("nan", "")
            else:
                # If "SID" column is missing, create it with default empty strings
                df["SID"] = ""

            # Replace any NaN with empty strings explicitly to avoid issues
            df["SID"] = df["SID"].fillna("")

            # Sheet 10: Rules-Renewal
            df = pd.read_excel(file1, sheet_name="Rules-Renewal")

            # Convert "Max Cycle" and "Period" columns to integers
            df["Max Cycle"] = pd.to_numeric(df["Max Cycle"], errors="coerce").astype("Int64")
            df["Period"] = pd.to_numeric(df["Period"], errors="coerce").astype("Int64")

            # Remove commas and decimals from "Amount" and truncate decimals
            if "Amount" in df.columns:
                # Remove commas, keep only numeric part, and truncate decimals
                df["Amount"] = (
                    df["Amount"]
                    .str.replace(",", "", regex=False)  # Remove commas
                    .str.split(".", n=1).str[0]        # Remove decimals by splitting
                )
                # Convert to integer
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").astype("Int64")
            else:
                df["Amount"] = None  # Handle cases where "Amount" column is missing

            # Save the modified DataFrame to the Excel sheet
            df.to_excel(writer, sheet_name="Rules-Renewal", index=False)


            # Sheet 11: Rules-GSI GRP Pack
            gsi_grp_pack_df = pd.DataFrame(
                {
                    "Ruleset ShortName": ["sample"],  # First row value
                    "GSI GRP Pack-Group ID": ["sample"],  # First row value
                    "Action": ["NO_CHANGE"],  # First row value
                }
            )
            gsi_grp_pack_df.to_excel(writer, sheet_name="Rules-GSI GRP Pack", index=False)

            # Sheet 12: Rules-Location Group
            location_group_df = pd.DataFrame(
                {
                    "Ruleset ShortName": ["sample"],
                    "Package Group": ["sample"],
                    "Microcluster ID": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            location_group_df.to_excel(writer, sheet_name="Rules-Location Group", index=False)

            # Sheet 13: Rebuy-Out
            rebuy_out_df = pd.DataFrame(
                {
                    "Target PO ID": ["sample"],
                    "Target Ruleset ShortName": ["sample"],
                    "Target MPP": ["sample"],
                    "Target Group": ["sample"],
                    "Service Type": ["sample"],
                    "Rebuy Price": ["sample"],
                    "Allow Rebuy": ["sample"],
                    "Rebuy Option": ["sample"],
                    "Product Family": ["sample"],
                    "Source PO ID": ["sample"],
                    "Source Ruleset ShortName": ["sample"],
                    "Source MPP": ["sample"],
                    "Source Group": ["sample"],
                    "Vice Versa Consent": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            rebuy_out_df.to_excel(writer, sheet_name="Rebuy-Out", index=False)

            # Sheet 14: Rebuy-Association
            rebuy_association_df = pd.DataFrame(
                {
                    "Target PO ID": ["sample"],
                    "Target Ruleset ShortName": ["sample"],
                    "Target MPP": ["sample"],
                    "Target Group": ["sample"],
                    "Service Type": ["sample"],
                    "Rebuy Price": ["sample"],
                    "Allow Rebuy": ["sample"],
                    "Rebuy Option": ["sample"],
                    "Product Family": ["sample"],
                    "Source PO ID": ["sample"],
                    "Source Ruleset ShortName": ["sample"],
                    "Source MPP": ["sample"],
                    "Source Group": ["sample"],
                    "Vice Versa Consent": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            rebuy_association_df.to_excel(writer, sheet_name="Rebuy-Association", index=False)

            # Sheet 15: Incompatibility
            incompatibility_df = pd.DataFrame(
                {
                    "ID": ["sample"],
                    "Target PO/RulesetShortName": ["sample"],
                    "Source Family": ["sample"],
                    "Source PO/RulesetShortName": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            incompatibility_df.to_excel(writer, sheet_name="Incompatibility", index=False)

            # Sheet 16: Library-Addon-Name
            df = pd.read_excel(file1, sheet_name="Rules-Library-Addon")
            df.to_excel(writer, sheet_name="Library-Addon-Name", index=False)

            # Sheet 17: Library-Addon-DA -later get it from DDM
            library_addon_da_df = pd.DataFrame(
                {
                    "Ruleset ShortName": ["sample"],
                    "PO ID": ["sample"],
                    "Quota Name": ["sample"],
                    "DA ID": ["sample"],
                    "Internal Description Bahasa": ["sample"],
                    "External Description Bahasa": ["sample"],
                    "Internal Description English": ["sample"],
                    "External Description English": ["sample"],
                    "Visibility": ["sample"],
                    "Custom": ["sample"],
                    "Feature": ["sample"],
                    "Initial Value": ["sample"],
                    "Unlimited Benefit Flag": ["sample"],
                    "Scenario": ["sample"],
                    "Attribute Name": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            library_addon_da_df.to_excel(writer, sheet_name="Library-Addon-DA", index=False)

            # Sheet 18: Library-Addon-UCUT
            library_addon_ucut_df = pd.DataFrame(
                {
                    "Ruleset ShortName": ["sample"],
                    "PO ID": ["sample"],
                    "Quota Name": ["sample"],
                    "UCUT ID": ["sample"],
                    "Internal Description Bahasa": ["sample"],
                    "External Description Bahasa": ["sample"],
                    "Internal Description English": ["sample"],
                    "External Description English": ["sample"],
                    "Visibility": ["sample"],
                    "Custom": ["sample"],
                    "Initial Value": ["sample"],
                    "Unlimited Benefit Flag": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            library_addon_ucut_df.to_excel(writer, sheet_name="Library-Addon-UCUT", index=False)

            # Sheet 19: Standalone - later get it from DDM
            standalone_df = pd.DataFrame(
                {
                    "Ruleset ShortName": ["sample"],
                    "PO ID": ["sample"],
                    "Scenarios": ["sample"],
                    "Type": ["sample"],
                    "ID": ["sample"],
                    "Value": ["sample"],
                    "UOM": ["sample"],
                    "Validity": ["sample"],
                    "Provision Payload Value": ["sample"],
                    "Payload Dependent Attribute": ["sample"],
                    "ACTION": ["sample"],
                    "Action": ["NO_CHANGE"],
                }
            )
            standalone_df.to_excel(writer, sheet_name="Standalone", index=False)

            # Sheet 20: Blacklist-Gift-Promocodes
            blacklist_gift_promocodes_df = pd.DataFrame(
                [{"Ruleset ShortName": "sample", "Coherence Key": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
            )
            blacklist_gift_promocodes_df.to_excel(writer, sheet_name="Blacklist-Gift-Promocodes", index=False)

            # Sheet 21: Blacklist-Promocodes
            blacklist_promocodes_df = pd.DataFrame(
                [{"PO ID": "sample", "Command/Keyword": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
            )
            blacklist_promocodes_df.to_excel(writer, sheet_name="Blacklist-Promocodes", index=False)

            # Sheet 22: MYIM3-UNREG
            myim3_unreg_df = pd.DataFrame(
                [
                    {
                        "Ruleset ShortName": "sample",
                        "Keyword": "sample",
                        "Shortcode": "sample",
                        "Unreg Flag": "sample",
                        "Buy Extra Flag": "sample",
                        "Action": "NO_CHANGE",
                    }
                ]
            )
            myim3_unreg_df.to_excel(writer, sheet_name="MYIM3-UNREG", index=False)

            # Sheet 23: ExtraPOConfig
            extrapoconfig_df = pd.DataFrame(
                [{"Ruleset ShortName": "sample", "Extra PO Keyword": "sample", "Action": "NO_CHANGE"}]
            )
            extrapoconfig_df.to_excel(writer, sheet_name="ExtraPOConfig", index=False)

            # Sheet 24: Keyword-Global-Variable
            keyword_global_variable_df = pd.DataFrame(
                [
                    {
                        "PO ID": "sample",
                        "Keyword": "sample",
                        "Global Variable Type": "sample",
                        "Value": "sample",
                        "Keyword Type": "sample",
                        "Action": "NO_CHANGE",
                    }
                ]
            )
            keyword_global_variable_df.to_excel(writer, sheet_name="Keyword-Global-Variable", index=False)

            # Sheet 25: UMB-Push-Category
            umb_push_category_df = pd.DataFrame(
                [
                    {
                        "Ruleset ShortName": "sample",
                        "Coherence Key": "sample",
                        "Group Category": "sample",
                        "Short Code": "sample",
                        "Show Unit": "sample",
                        "Action": "NO_CHANGE",
                    }
                ]
            )
            umb_push_category_df.to_excel(writer, sheet_name="UMB-Push-Category", index=False)

            # Sheet 26: Avatar-Channel
            avatar_channel_df = pd.DataFrame(
                [
                    {
                        "PO ID": "sample",
                        "Ruleset ShortName": "sample",
                        "Keyword": "sample",
                        "Commercial Name": "sample",
                        "Short Code": "sample",
                        "PVR ID": "sample",
                        "Price": "sample",
                        "Action": "NO_CHANGE",
                    }
                ]
            )
            avatar_channel_df.to_excel(writer, sheet_name="Avatar-Channel", index=False)

            # Sheet 27: Dormant-Config
            dormant_config_df = pd.DataFrame(
                [{"Ruleset ShortName": "sample", "Keyword": "sample", "Short Code": "sample", "Pvr": "sample", "Action": "NO_CHANGE"}]
            )
            dormant_config_df.to_excel(writer, sheet_name="Dormant-Config", index=False)

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
