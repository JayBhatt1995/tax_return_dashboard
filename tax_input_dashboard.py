import streamlit as st
import pandas as pd
import datetime
import os

# Load data from Excel
excel_path = "Billable hours vs. non-billable hours.xlsx"
table_df = pd.read_excel(excel_path, sheet_name="Table")
input_template = pd.read_excel(excel_path, sheet_name="Input field")

# Dropdown values
return_types = table_df['Return Type'].dropna().unique().tolist()
service_types = table_df['Service type'].dropna().unique().tolist()
office_locations = table_df['Office Location'].dropna().unique().tolist()
complexity_levels = table_df['Complexity Level'].dropna().unique().tolist()
if_btr_options = table_df['IF BTR '].dropna().unique().tolist()
tax_associates = table_df[table_df['Designation '] == 'Tax Associate']['Employee Name'].dropna().unique().tolist()
tax_reviewers = table_df[table_df['Designation '] == 'Tax Reviewer']['Employee Name'].dropna().unique().tolist()

# App title
st.markdown("<h1 style='color: rgb(23, 45, 100);'>VGSL</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='color: rgb(23, 45, 100);'>Tax Return Automation Dashboard</h2>", unsafe_allow_html=True)

# Data path
data_file = "input_data.xlsx"

# Load and clean existing data
if os.path.exists(data_file):
    df_existing = pd.read_excel(data_file)
    df_existing.columns = df_existing.columns.str.strip()
    df_existing = df_existing.loc[:, ~df_existing.columns.duplicated()]
else:
    df_existing = pd.DataFrame()

# Step 1: Ask for Client ID
st.markdown("### Step 1: Enter or Search by Client ID")
client_id_input = st.text_input("Enter Client ID to Edit (or Create New Entry)").strip()

# Look up existing data
existing_data = None
if client_id_input and not df_existing.empty and "Client ID" in df_existing.columns:
    match = df_existing[df_existing["Client ID"].astype(str) == client_id_input]
    if not match.empty:
        existing_data = match.iloc[0]

# Helper functions
def get_val(key, default=""):
    return existing_data[key] if existing_data is not None and key in existing_data else default

def get_date_val(key):
    val = get_val(key)
    try:
        parsed = pd.to_datetime(val)
        if pd.isna(parsed):
            raise ValueError("Invalid date (NaT)")
        return parsed.date()
    except:
        return datetime.date.today()

# Step 2: Form to view/edit
if client_id_input:
    with st.form("tax_form"):
        client_name = st.text_input("Client Name", value=get_val("Client Name"))
        return_type = st.selectbox("Return Type", return_types, index=return_types.index(get_val("Return Type")) if get_val("Return Type") in return_types else 0)
        service_type = st.selectbox("Service Type", service_types, index=service_types.index(get_val("Service Type")) if get_val("Service Type") in service_types else 0)
        office_location = st.selectbox("Office Location", office_locations, index=office_locations.index(get_val("Office Location")) if get_val("Office Location") in office_locations else 0)
        new_resubmitted = st.selectbox("New/Resubmitted", ['New', 'Resubmitted'], index=['New', 'Resubmitted'].index(get_val("New/Resubmitted")) if get_val("New/Resubmitted") in ['New', 'Resubmitted'] else 0)
        Pages = st.text_input("Pages", value=get_val("Pages"))
        Return_for_Preparation_or_Review = st.text_input("Return for Preparation or Review", value=get_val("Return for Preparation or Review"))
        Return_pull_back_date = st.date_input("Return pull back date", value=get_date_val("Return pull back date"))
        preparer = st.selectbox("Preparer Name", tax_associates, index=tax_associates.index(get_val("Preparer Name")) if get_val("Preparer Name") in tax_associates else 0)
        Date_of_Allocation = st.date_input("Date of Allocation", value=get_date_val("Date of Allocation"))
        start_date_1 = st.date_input("Start Date 1", value=get_date_val("start date 1"))
        end_date_1 = st.date_input("End Date 1", value=get_date_val("end date 1"))
        total_time_spent = st.text_input("Total Time Spent", value=get_val("total time spent"))
        reviewer = st.selectbox("Reviewer Name", tax_reviewers, index=tax_reviewers.index(get_val("Reviewer Name")) if get_val("Reviewer Name") in tax_reviewers else 0)
        if_btr = st.selectbox("IF BTR", if_btr_options, index=if_btr_options.index(get_val("IF BTR")) if get_val("IF BTR") in if_btr_options else 0)
        start_date_2 = st.date_input("Start Date 2", value=get_date_val("start date 2"))
        end_date_2 = st.date_input("End Date 2", value=get_date_val("end date 2"))
        total_time = st.text_input("Total Time", value=get_val("total time"))
        complexity_level = st.selectbox("Complexity Level", complexity_levels, index=complexity_levels.index(get_val("Complexity Level")) if get_val("Complexity Level") in complexity_levels else 0)
        Remarks = st.text_input("Remarks", value=get_val("Remarks"))
        Total_return_time = st.text_input("Total Return Time", value=get_val("Total return time"))
        Return_submission_date = st.date_input("Return Submission Date", value=get_date_val("Return submission date"))

        submit = st.form_submit_button("Submit Entry")

    if submit:
        new_entry = {
            "Client ID": client_id_input,
            "Client Name": client_name,
            "Return Type": return_type,
            "Service Type": service_type,
            "Office Location": office_location,
            "New/Resubmitted": new_resubmitted,
            "Pages": Pages,
            "Return for Preparation or Review": Return_for_Preparation_or_Review,
            "Return pull back date": Return_pull_back_date,
            "Preparer Name": preparer,
            "Date of Allocation": Date_of_Allocation,
            "start date 1": start_date_1,
            "end date 1": end_date_1,
            "total time spent": total_time_spent,
            "Reviewer Name": reviewer,
            "IF BTR": if_btr,
            "start date 2": start_date_2,
            "end date 2": end_date_2,
            "total time": total_time,
            "Complexity Level": complexity_level,
            "Remarks": Remarks,
            "Total return time": Total_return_time,
            "Return submission date": Return_submission_date
        }

        # Remove previous if same client ID
        if "Client ID" in df_existing.columns:
            df_existing = df_existing[df_existing["Client ID"].astype(str) != client_id_input]

        # Append and save
        df_updated = pd.concat([df_existing, pd.DataFrame([new_entry])], ignore_index=True)
        df_updated.to_excel(data_file, index=False)
        st.success("Entry saved successfully!")

# Step 3: Display report
if st.button("Display & Download Report"):
    if os.path.exists(data_file):
        df_final = pd.read_excel(data_file)
        df_final.columns = df_final.columns.str.strip()
        df_final = df_final.loc[:, ~df_final.columns.duplicated()]
        df_final.drop_duplicates(subset="Client ID", keep="last", inplace=True)
        st.dataframe(df_final)
        st.download_button("Download as Excel Report", df_final.to_csv(index=False), "final_report.csv")
    else:
        st.warning("No records found.")
