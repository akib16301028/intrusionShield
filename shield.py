import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications
import os  # For file path operations
from io import BytesIO

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    if pd.notnull(site_name):
        parts = site_name.split('_')
        return parts[0] if len(parts) > 1 else site_name
    return site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    # Debug: Print columns in RMS and Alarms DataFrames
    st.write("Columns in RMS DataFrame:", rms_df.columns.tolist())
    st.write("Columns in Alarms DataFrame:", alarms_df.columns.tolist())

    # Ensure required columns exist
    required_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Alarm Time']
    for col in required_columns:
        if col not in alarms_df.columns:
            st.error(f"Column '{col}' not found in Alarms DataFrame.")
            return pd.DataFrame()  # Return empty DataFrame if columns are missing

    # Prepare Alarms DataFrame
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    # Prepare RMS DataFrame
    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    # Ensure columns exist in RMS DataFrame
    for col in rms_columns:
        if col not in rms_df.columns:
            st.error(f"Column '{col}' not found in RMS DataFrame.")
            return pd.DataFrame()  # Return empty DataFrame if columns are missing

    # Merge DataFrames
    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    st.write("Merged DataFrame:")
    st.write(merged_df.head())
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    # Debug: Print columns in Site Access DataFrame
    st.write("Columns in Site Access DataFrame:", site_access_df.columns.tolist())

    # Ensure required columns exist
    if 'SiteName' not in site_access_df.columns:
        st.error("Column 'SiteName' not found in Site Access DataFrame.")
        return pd.DataFrame()  # Return empty DataFrame if columns are missing

    # Extract site names
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    st.write("Site Access DataFrame with Extracted Site Names:")
    st.write(site_access_df.head())

    # Merge with merged RMS/Alarms DataFrame
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    st.write("Merged Comparison DataFrame:")
    st.write(merged_comparison_df.head())

    # Find mismatches
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')  # Replace NaT with Not Closed
    st.write("Mismatched Sites DataFrame:")
    st.write(mismatches_df)
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

# Function to display grouped data by Cluster and Zone in a table
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
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            display_df['Site Alias'] = display_df['Site Alias'].where(display_df['Site Alias'] != display_df['Site Alias'].shift())
            display_df = display_df.fillna('')
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

# Function to send Telegram notification
def send_telegram_notification(message, bot_token, chat_id):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "Markdown"  # Use Markdown for plain text
    }
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Streamlit app
st.title('🛡️IntrusionShield🛡️')

# File uploaders
site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Door Open Alarms Data", type=["xlsx"])

# Session state for filters
if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

# Process files if uploaded
if site_access_file and rms_file and current_alarms_file:
    # Load data
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    # Debug: Display first few rows of each DataFrame
    st.write("### Debugging: First Few Rows of Each File")
    st.write("Site Access Data:")
    st.write(site_access_df.head())
    st.write("RMS Data (All Door Open Alarms till now):")
    st.write(rms_df.head())
    st.write("Current Alarms Data:")
    st.write(current_alarms_df.head())

    # Merge RMS and Current Alarms DataFrames
    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Find mismatches
    if not merged_rms_alarms_df.empty:
        mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
        if not mismatches_df.empty:
            st.write("### Mismatched Sites Found")
            display_grouped_data(mismatches_df, "Mismatched Sites Grouped by Cluster and Zone")
        else:
            st.write("No mismatched sites found.")
    else:
        st.error("Merged DataFrame is empty. Check the input files and column names.")

    # Find matched sites
    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)
    if not matched_df.empty:
        st.write("### Matched Sites with Status")
        display_matched_sites(matched_df)
    else:
        st.write("No matched sites found.")

# Sidebar options
st.sidebar.title("Options")
st.sidebar.write("### Update Zone Concern")
user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

if os.path.exists(user_file_path):
    user_df = pd.read_excel(user_file_path)
    if "Zone" in user_df.columns and "Name" in user_df.columns:
        zone_list = user_df['Zone'].unique()
        selected_zone = st.sidebar.selectbox("Select Zone", options=zone_list)
        current_name = user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0]
        new_name = st.sidebar.text_input("Update Name", value=current_name)

        if st.sidebar.button("🔄Update Concern"):
            user_df.loc[user_df['Zone'] == selected_zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            st.sidebar.success("Zone concern updated successfully!")
    else:
        st.sidebar.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
else:
    st.sidebar.error("USER NAME.xlsx file not found in the repository.")

# Telegram Notification Option
if st.sidebar.button("💬 Send Notification"):
    if not mismatches_df.empty:
        bot_token = "YOUR_BOT_TOKEN"
        chat_id = "YOUR_CHAT_ID"
        message = "❗Door Open Notification❗\n\n"
        for _, row in mismatches_df.iterrows():
            message += f"🚩 {row['Site Alias']} ({row['Zone']})\n"
            message += f"  • Start Time: {row['Start Time']}\n"
            message += f"  • End Time: {row['End Time']}\n\n"
        if send_telegram_notification(message, bot_token, chat_id):
            st.sidebar.success("Notification sent successfully!")
        else:
            st.sidebar.error("Failed to send notification.")
    else:
        st.sidebar.write("No mismatched sites to notify.")

# Download Option
if site_access_file and rms_file and current_alarms_file:
    timestamp = datetime.now().strftime("%d%m%y%H%M%S")
    file_name = f"UnauthorizedAccess_{timestamp}.xlsx"
    excel_data = BytesIO()
    with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
        mismatches_df.to_excel(writer, sheet_name='Mismatched Sites', index=False)
        matched_df.to_excel(writer, sheet_name='Matched Sites', index=False)
    excel_data.seek(0)
    st.sidebar.download_button(
        label="📂 Download Data",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
