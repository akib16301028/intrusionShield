import pandas as pd
import streamlit as st
from datetime import datetime
import requests
import os
from io import BytesIO

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT
    rms_columns = ['Site', 'Site Alias ', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')
    return mismatches_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')
    matched_df['Status'] = matched_df.apply(lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

# FIXED: Function to display grouped data by Cluster and Zone in a table
def display_grouped_data(grouped_df, title):
    st.write(title)
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            # Handle both column name variations
            alias_col = 'Site Alias' if 'Site Alias' in zone_df.columns else 'Site Alias '
            display_df = zone_df[[alias_col, 'Start Time', 'End Time']].copy()
            display_df.columns = ['Site Alias', 'Start Time', 'End Time']
            
            # Reset index to avoid index-related errors
            display_df = display_df.reset_index(drop=True)
            
            # Fill NA values and ensure proper display
            display_df['Site Alias'] = display_df['Site Alias'].fillna('Unknown')
            
            # Only show Site Alias when it changes - using shift() which is more reliable
            display_df['Site Alias'] = display_df['Site Alias'].where(
                display_df['Site Alias'] != display_df['Site Alias'].shift(), 
                ''
            )
            
            st.table(display_df)
        st.markdown("---")

# Function to display matched sites with status
def display_matched_sites(matched_df):
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
    def highlight_status(status):
        return color_map.get(status, '')
    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])
    st.write("Matched Sites with Status:")
    st.dataframe(styled_df)

# FIXED: Function to send Telegram notification
def send_telegram_notification(mismatches_df, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
            bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
            chat_id = "-4625672098"

            zones = mismatches_df['Zone'].unique()
            
            for zone in zones:
                zone_df = mismatches_df[mismatches_df['Zone'] == zone]
                message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                
                # Handle both column name variations
                alias_col = 'Site Alias' if 'Site Alias' in zone_df.columns else 'Site Alias '
                
                for site_alias in zone_df[alias_col].dropna().unique():
                    site_df = zone_df[zone_df[alias_col] == site_alias]
                    message += f"✔ {site_alias}\n"
                    for _, row in site_df.iterrows():
                        end_time = row['End Time'] if pd.notna(row['End Time']) else "Not Closed"
                        message += f"  • Start: {row['Start Time']} | End: {end_time}\n"
                    message += "\n"

                if zone in zone_to_name:
                    escaped_name = zone_to_name[zone].replace("_", "\\_")
                    message += f"**@{escaped_name}**, no Site Access Request found for these alarms.\n"

                response = requests.post(
                    f"https://api.telegram.org/bot{bot_token}/sendMessage",
                    json={"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
                )

# Function to update the user name for a specific zone
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
        return False, "The USER NAME.xlsx file must have 'Zone' and 'Name' columns."
    return False, "USER NAME.xlsx file not found in the repository."

# Function to convert dataframes into an Excel file with multiple sheets
@st.cache_data
def convert_df_to_excel_with_sheets(unmatched_df, rms_df, current_alarms_df, site_access_df):
    filtered_unmatched_df = unmatched_df[['Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered_unmatched_df.to_excel(writer, index=False, sheet_name='Unmatched Data')
        rms_df.to_excel(writer, index=False, sheet_name='RMS Data')
        current_alarms_df.to_excel(writer, index=False, sheet_name='Current Alarms')
        site_access_df.to_excel(writer, index=False, sheet_name='Site Access Data')
        workbook = writer.book
        for sheet_name, df in [
            ('Unmatched Data', filtered_unmatched_df),
            ('RMS Data', rms_df),
            ('Current Alarms', current_alarms_df),
            ('Site Access Data', site_access_df)
        ]:
            worksheet = writer.sheets[sheet_name]
            for i, column in enumerate(df.columns):
                max_len = max(df[column].astype(str).map(len).max(), len(column)) + 2
                worksheet.set_column(i, i, max_len)
            if sheet_name == 'Unmatched Data':
                table_range = f'A1:E{len(filtered_unmatched_df) + 1}'
                worksheet.add_table(table_range, {
                    'columns': [{'header': col} for col in filtered_unmatched_df.columns],
                    'style': 'Table Style Medium 9',
                })
    return output.getvalue()

# Streamlit app
st.title('🛡️IntrusionShield🛡️')

site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Door Open Alarms Data", type=["xlsx"])

if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    selected_date = st.date_input("Select Date", value=st.session_state.filter_date)
    selected_time = st.time_input("Select Time", value=st.session_state.filter_time)

    if st.button("Clear Filters"):
        st.session_state.filter_date = datetime.now().date()
        st.session_state.filter_time = datetime.now().time()
        st.session_state.status_filter = "All"

    if selected_date != st.session_state.filter_date:
        st.session_state.filter_date = selected_date
    if selected_time != st.session_state.filter_time:
        st.session_state.filter_time = selected_time

    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)

    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)
    status_filter_condition = matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True
    time_filter_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)
    filtered_matched_df = matched_df[status_filter_condition & time_filter_condition]

    status_filter = st.selectbox("SA-Request Valid/Expired", options=["All", "Valid", "Expired"], index=0)
    if status_filter != st.session_state.status_filter:
        st.session_state.status_filter = status_filter

    if not filtered_mismatches_df.empty:
        st.write(f"Mismatched Sites (After {filter_datetime}) grouped by Cluster and Zone:")
        display_grouped_data(filtered_mismatches_df, "Filtered Mismatched Sites")
    else:
        st.write(f"No mismatches found after {filter_datetime}. Showing all mismatched sites.")
        display_grouped_data(mismatches_df, "All Mismatched Sites")

    display_matched_sites(filtered_matched_df)

# Sidebar options
st.sidebar.title("Options")
user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

if os.path.exists(user_file_path):
    user_df = pd.read_excel(user_file_path)
    if "Zone" in user_df.columns and "Name" in user_df.columns:
        zone_list = user_df['Zone'].unique()
        selected_zone = st.sidebar.selectbox("Select Zone", options=zone_list)
        current_name = user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0]
        new_name = st.sidebar.text_input("Update Name", value=current_name)
        if st.sidebar.button("🔄Update Concern"):
            success, message = update_zone_user(selected_zone, new_name, user_file_path)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.error(message)
    else:
        st.sidebar.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
else:
    st.sidebar.error("USER NAME.xlsx file not found in the repository.")

# Download Option
if site_access_file and rms_file and current_alarms_file:
    timestamp = datetime.now().strftime("%d%m%y%H%M%S")
    file_name = f"UnauthorizedAccess_{timestamp}.xlsx"
    excel_data = convert_df_to_excel_with_sheets(mismatches_df, rms_df, current_alarms_df, site_access_df)
    st.sidebar.download_button(
        label="📂 Download Data",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.sidebar.write("Please upload all required files to enable data download.")

# Telegram Notification Option
if st.sidebar.button("💬 Send Notification"):
    send_telegram_notification(filtered_mismatches_df if not filtered_mismatches_df.empty else mismatches_df, user_file_path)
