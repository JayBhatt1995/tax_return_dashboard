import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import os

# Load data from Excel
excel_path = "Billable hours vs. non-billable hours.xlsx"
table_df = pd.read_excel(excel_path, sheet_name="Table")
input_template = pd.read_excel(excel_path, sheet_name="Input field")

# Define dropdown options
return_types = table_df['Return Type'].dropna().unique().tolist()
service_types = table_df['Service type'].dropna().unique().tolist()
office_locations = table_df['Office Location'].dropna().unique().tolist()
complexity_levels = table_df['Complexity Level'].dropna().unique().tolist()
if_btr_options = table_df['IF BTR '].dropna().unique().tolist()
tax_associates = table_df['Employee Name'][table_df['Designation '] == 'Tax Associate'].dropna().unique().tolist()
tax_reviewers = table_df['Employee Name'][table_df['Designation '] == 'Tax Reviewer'].dropna().unique().tolist()

st.markdown(
    "<h1 style='color: rgb(23, 45, 100);'>VGSL</h1>",
    unsafe_allow_html=True
)

st.markdown(
    "<h2 style='color: rgb(23, 45, 100);'>Tax Return Automation Dashboard</h2>",
    unsafe_allow_html=True
)

# User Form Input
with st.form("tax_form"):
    return_received_date = st.date_input("Select a date", value=datetime.date.today())
    client_id = st.text_input("Client ID")
    client_name = st.text_input("Client Name")                                     
    return_type = st.selectbox("Return Type", return_types)
    service_type = st.selectbox("Service Type", service_types)
    office_location = st.selectbox("Office Location", office_locations)
    new_resubmitted = st.selectbox("New/Resubmitted", ['New', 'Resubmitted'])
    Pages = st.text_input("Pages")
    Return_for_Preparation_or_Review = st.text_input("Return for Preparation or Review")
    Return_pull_back_date =  st.date_input("Return pull back date", value=datetime.date.today())
    preparer = st.selectbox("Preparer Name", tax_associates)
    Date_of_Allocation = st.date_input("Date of Allocation", value=datetime.date.today())
    start_date_1  = st.date_input("start date 1", value=datetime.date.today(), key="start_date_input_1")
    end_date_1  = st.date_input("end date 1", value=datetime.date.today(), key="end_date_input_1")
    total_time_spent = st.text_input("total time spent")
    reviewer = st.selectbox("Reviewer Name", tax_reviewers)
    if_btr = st.selectbox("IF BTR ", if_btr_options)
    start_date_2  = st.date_input("start date 2", value=datetime.date.today(), key="start_date_input_2")
    end_date_2  = st.date_input("end date 2", value=datetime.date.today(), key="end_date_input_2")
    total_time = st.text_input("total time")
    complexity_level = st.selectbox("Complexity Level", complexity_levels)
    Remarks = st.text_input("Remarks")
    Total_return_time = st.text_input("Total return time")
    Return_submission_date = st.date_input("Return submission date", value=datetime.date.today())

    submit = st.form_submit_button("Submit Entry")

if submit:
    new_entry = {        
        "Client ID": client_id,
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
        "start date 1 ": start_date_1,
        "end date 1 ": end_date_1,
        "total time spent": total_time_spent,
        "Reviewer Name": reviewer,
        "IF BTR ": if_btr,
        "start date 2": start_date_2,
        "end date 2": end_date_2,
        "total time": total_time,        
        "Complexity Level": complexity_level,
        "Remarks": Remarks,
        "Total return time": Total_return_time,
        "Return submission date": Return_submission_date        
    }

    try:
        df_existing = pd.read_excel("input_data.xlsx")
    except:
        df_existing = pd.DataFrame()

    df_updated = pd.concat([df_existing, pd.DataFrame([new_entry])], ignore_index=True)
    df_updated.to_excel('input_data.xlsx', index=False)
    st.success("Entry saved successfully!")

# Show consolidated data
if st.button("Display & Download Report"):
    df_final = pd.read_excel('input_data.xlsx')
    df_final.drop_duplicates(subset='Client ID', keep='last', inplace=True)
    st.dataframe(df_final)
    st.download_button("Download as Excel Report", df_final.to_csv(index=False), "final_report.csv")