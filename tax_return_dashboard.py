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
tax_associates = table_df[table_df['Designation '] == 'Tax Associate']['Employee Name'].dropna().unique().tolist()
tax_reviewers = table_df[table_df['Designation '] == 'Tax Reviewer']['Employee Name'].dropna().unique().tolist()

# Updated IF BTR dropdown options
if_btr_options = ['TB Import','K1','M1 /M3','Entries','Access Input','1065','1120','1120S','990']
if_1040_1041_options = ['1040', '1041']

# App title
st.markdown("<h1 style='color: rgb(23, 45, 100);'>VGSL</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='color: rgb(23, 45, 100);'>Tax Return Dashboard</h2>", unsafe_allow_html=True)

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
        return None

def get_int_val(key):
    val = get_val(key)
    try:
        return int(round(float(val)))
    except:
        return 0

# Step 2: Form to view/edit
if client_id_input:
    try:
        with st.form("tax_form"):
            binder_id = st.text_input("Binder ID", value=get_val("Binder ID"))
            client_name = st.text_input("Client Name", value=get_val("Client Name"))
            return_type = st.selectbox("Return Type", return_types, index=return_types.index(get_val("Return Type")) if get_val("Return Type") in return_types else 0)
            service_type = st.selectbox("Service Type", service_types, index=service_types.index(get_val("Service Type")) if get_val("Service Type") in service_types else 0)
            office_location = st.selectbox("Office Location", office_locations, index=office_locations.index(get_val("Office Location")) if get_val("Office Location") in office_locations else 0)
            new_resubmitted = st.selectbox("New/Resubmitted", ['New', 'Resubmitted'], index=['New', 'Resubmitted'].index(get_val("New/Resubmitted")) if get_val("New/Resubmitted") in ['New', 'Resubmitted'] else 0)
            Pages = st.text_input("Pages", value=get_val("Pages"))
            Return_for_Preparation_or_Review = st.selectbox("Return for Preparation or Review", ['Preparation', 'Review'], index=['Preparation', 'Review'].index(get_val("Return for Preparation or Review")) if get_val("Return for Preparation or Review") in ['Preparation', 'Review'] else 0)

            pull_back_option = st.radio("Return pull back date", ["Not Applicable", "Select Date"], index=0)
            Return_pull_back_date = None
            if pull_back_option == "Select Date":
                Return_pull_back_date = st.date_input("Select pull back date", value=get_date_val("Return pull back date") or datetime.date.today())

            preparer = st.selectbox("Preparer Name", tax_associates, index=tax_associates.index(get_val("Preparer Name")) if get_val("Preparer Name") in tax_associates else 0)
            Date_of_Allocation = st.date_input("Date of Allocation", value=get_date_val("Date of Allocation") or datetime.date.today())

            start_date_1 = st.date_input("Start Date 1", value=get_date_val("start date 1") or datetime.date.today())
            end_date_1 = st.date_input("End Date 1", value=get_date_val("end date 1") or datetime.date.today(), min_value=start_date_1)
            start_date_2 = st.date_input("Start Date 2", value=get_date_val("start date 2") or datetime.date.today(), min_value=end_date_1)
            end_date_2 = st.date_input("End Date 2", value=get_date_val("end date 2") or datetime.date.today(), min_value=start_date_2)
            Return_submission_date = st.date_input("Return Submission Date", value=get_date_val("Return submission date") or datetime.date.today(), min_value=end_date_2)

            total_time_spent = st.number_input("Total Time Spent (rounded)", value=get_int_val("total time spent"), step=1)
            reviewer = st.selectbox("Reviewer Name", tax_reviewers, index=tax_reviewers.index(get_val("Reviewer Name")) if get_val("Reviewer Name") in tax_reviewers else 0)
            if_btr = st.multiselect("IF BTR", if_btr_options, default=[val for val in if_btr_options if val in str(get_val("IF BTR"))])
            if_1040_1041 = st.selectbox("IF 1040/1041", if_1040_1041_options, index=if_1040_1041_options.index(get_val("IF 1040/1041")) if get_val("IF 1040/1041") in if_1040_1041_options else 0)
            total_time = st.number_input("Total Time (rounded)", value=get_int_val("total time"), step=1)

            Total_return_time = total_time + total_time_spent
            st.number_input("Total Return Time (rounded)", value=Total_return_time, step=1, disabled=True)

            complexity_level = st.selectbox("Complexity Level", complexity_levels, index=complexity_levels.index(get_val("Complexity Level")) if get_val("Complexity Level") in complexity_levels else 0)
            Remarks = st.text_input("Remarks", value=get_val("Remarks"))

            submit = st.form_submit_button("Submit Entry")

        if submit:
            new_entry = {
                "Client ID": client_id_input,
                "Binder ID": binder_id,
                "Client Name": client_name,
                "Return Type": return_type,
                "Service Type": service_type,
                "Office Location": office_location,
                "New/Resubmitted": new_resubmitted,
                "Pages": Pages,
                "Return for Preparation or Review": Return_for_Preparation_or_Review,
                "Return pull back date": Return_pull_back_date if Return_pull_back_date else "",
                "Preparer Name": preparer,
                "Date of Allocation": Date_of_Allocation,
                "start date 1": start_date_1,
                "end date 1": end_date_1,
                "start date 2": start_date_2,
                "end date 2": end_date_2,
                "Return submission date": Return_submission_date,
                "total time spent": total_time_spent,
                "Reviewer Name": reviewer,
                "IF BTR": ", ".join(if_btr),
                "IF 1040/1041": if_1040_1041,
                "total time": total_time,
                "Complexity Level": complexity_level,
                "Remarks": Remarks,
                "Total return time": Total_return_time
            }

            if "Client ID" in df_existing.columns:
                df_existing = df_existing[df_existing["Client ID"].astype(str) != client_id_input]

            df_updated = pd.concat([df_existing, pd.DataFrame([new_entry])], ignore_index=True)
            df_updated.to_excel(data_file, index=False)
            st.success("Entry saved successfully!")
    except Exception as e:
        st.error(f"Something went wrong: {e}")

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
