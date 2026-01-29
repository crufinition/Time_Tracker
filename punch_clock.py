import streamlit as st
import pandas as pd
from datetime import datetime as dt
import calendar

# --- SETUP & DATA LOADING ---

def initialize_files(year=dt.today().year):
    try:
        pd.read_csv("員工清單.csv")
    except FileNotFoundError:
        pd.DataFrame(columns=['Name', 'Start Date', 'End Date']).to_csv("員工清單.csv", index=False)
    
    try:
        pd.read_csv(f"{year}打卡紀錄.csv")
    except FileNotFoundError:
        pd.DataFrame(columns=['Date', 'Employee', 'Hours']).to_csv(f"{year}打卡紀錄.csv", index=False)

initialize_files()

st.set_page_config(page_title="打卡紀錄", layout="wide")

def load_data(year=dt.today().year):
    # Load or create Employee Roster
    try:
        emp_df = pd.read_csv("員工清單.csv", parse_dates=['Start Date', 'End Date'])
    except FileNotFoundError:
        emp_df = pd.DataFrame(columns=['Name', 'Start Date', 'End Date'])
    
    # Load or create Time Logs
    try:
        time_df = pd.read_csv(f"{year}打卡紀錄.csv", parse_dates=['Date'])
    except FileNotFoundError:
        pd.DataFrame(columns=['Date', 'Employee', 'Hours']).to_csv(f"{year}打卡紀錄.csv", index=False)
        time_df = pd.DataFrame(columns=['Date', 'Employee', 'Hours'])
        
    return emp_df, time_df

emp_df, time_df = load_data()

# --- SIDEBAR FILTERS ---

st.sidebar.markdown("---")
st.sidebar.subheader("Select Period")

# 1. Year Entry & Month Pull-down
selected_year = st.sidebar.number_input("Year", min_value=2020, max_value=2099, value=dt.now().year)
month_names = ["January", "February", "March", "April", "May", "June", 
               "July", "August", "September", "October", "November", "December"]
selected_month_name = st.sidebar.selectbox("Month", month_names, index=dt.now().month - 1)
selected_month = month_names.index(selected_month_name) + 1

# --- DATA LOADING ---
# We force parse_dates to ensure they aren't treated as strings
emp_df, time_df = load_data()
st.sidebar.header("Navigation & Filters")
page = st.sidebar.radio("Go to:", ["輸入打卡紀錄", "編輯員工列表"])

# --- LOGIC: Filter Active Employees ---
def is_active(row, year, month):
    # Start date must be before or during the selected month
    start_match = row['Start Date'] <= dt(year, month, 1)
    assert False #最後編輯位置
    
    # End date must be empty (NaT) or after the start of the selected month
    if pd.isna(row['End Date']):
        end_match = True
    else:
        end_match = row['End Date'] >= dt(year, month, 1)
        
    return start_match and end_match

active_emps = emp_df[emp_df.apply(lambda r: is_active(r, selected_year, selected_month), axis=1)]['Name'].tolist()
# --- PAGE 1: TIME TRACKING (EXCEL-STYLE) ---
if page == "輸入打卡紀錄":
    st.title(f"📅 Timesheet: {selected_month_name} {selected_year}")
    
    if not active_emps:
        st.warning("No active employees found for this period.")
    else:
        selected_emp = st.selectbox("Select Employee", active_emps)

        # 1. Create a template for the FULL month
        last_day = calendar.monthrange(selected_year, selected_month)[1]
        all_days = pd.date_range(start=f"{selected_year}-{selected_month}-01", 
                                 end=f"{selected_year}-{selected_month}-{last_day}")
        template_df = pd.DataFrame({'Date': all_days})

        # 2. Merge with existing data
        existing_logs = time_df[time_df['Employee'] == selected_emp]
        # Merge template with existing logs to preserve already entered hours
        display_df = pd.merge(template_df, existing_logs[['Date', 'Hours']], on='Date', how='left')
        display_df = display_df.fillna(0) # Default hours to 0

        st.info(f"Filling hours for **{selected_emp}**. Scroll down to see all days.")

        # 3. Interactive Data Editor (Excel Style)
        edited_df = st.data_editor(
            display_df,
            column_config={
                "Date": st.column_config.DateColumn("Date", format="ddd, MMM DD", disabled=True),
                "Hours": st.column_config.NumberColumn("Hours Worked", min_value=0.0, max_value=24.0, step=0.5, format="%.1f")
            },
            hide_index=True,
            width='stretch',
            num_rows="fixed" # Prevents users from adding/deleting rows
        )

        if st.button("Save Monthly Records"):
            # Prepare data to save: Add employee name back and filter out zeros (optional)
            save_data = edited_df.copy()
            save_data['Employee'] = selected_emp
            
            # Remove old entries for this employee/month from master file
            mask = (time_df['Employee'] == selected_emp) & \
                   (time_df['Date'].dt.month == selected_month) & \
                   (time_df['Date'].dt.year == selected_year)
            
            final_df = pd.concat([time_df[~mask], save_data], ignore_index=True)
            final_df.to_csv("time_logs.csv", index=False)
            st.success(f"Successfully updated records for {selected_emp}!")
# --- PAGE 2: MANAGE EMPLOYEES ---
else:
    st.title("編輯員工列表")
    st.info("Update start/end dates here. Leave 'End Date' blank for current employees.")
    
    edited_emp_df = st.data_editor(emp_df, num_rows="dynamic", width='stretch')
    
    if st.button("Save Roster Changes"):
        edited_emp_df.to_csv("員工清單.csv", index=False)
        st.success("Roster updated!")
        st.rerun()