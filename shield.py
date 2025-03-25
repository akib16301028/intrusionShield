import pandas as pd
import streamlit as st
from datetime import datetime
import requests
import os
from io import BytesIO

# Helper functions
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

def merge_rms_alarms(rms_df, alarms_df):
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT
    return pd.concat([
        rms_df[['Site', 'Site Alias ', 'Zone', 'Cluster', 'Start Time', 'End Time']],
        alarms_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]
    ], ignore_index=True)

def find_mismatches(site_access_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged = pd.merge(alarms_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches = merged[merged['_merge'] == 'left_only']
    mismatches['End Time'] = mismatches['End Time'].fillna('Not Closed')
    return mismatches

def find_matched_sites(site_access_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched = pd.merge(site_access_df, alarms_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched['StartDate'] = pd.to_datetime(matched['StartDate'], errors='coerce')
    matched['EndDate'] = pd.to_datetime(matched['EndDate'], errors='coerce')
    matched['Start Time'] = pd.to_datetime(matched['Start Time'], errors='coerce')
    matched['End Time'] = pd.to_datetime(matched['End Time'], errors='coerce')
    matched['Status'] = matched.apply(lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched

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
            
            # Handle both 'Site Alias' and 'Site Alias ' columns
            alias_col = 'Site Alias' if 'Site Alias' in zone_df.columns else 'Site Alias '
            display_df = zone_df[[alias_col, 'Start Time', 'End Time']].copy()
            display_df.columns = ['Site Alias', 'Start Time', 'End Time']  # Standardize column names
            
            # Remove NaN values and empty strings
            display_df['Site Alias'] = display_df['Site Alias'].fillna('Unknown').replace('', 'Unknown')
            
            # Only show Site Alias when it changes
            for i in range(1, len(display_df)):
                if display_df.at[i, 'Site Alias'] == display_df.at[i-1, 'Site Alias']:
                    display_df.at[i, 'Site Alias'] = ''
            
            st.table(display_df.reset_index(drop=True))
        st.markdown("---")

def send_telegram_notification(mismatches_df, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
            bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
            chat_id = "-4625672098"

            for zone in mismatches_df['Zone'].unique():
                zone_df = mismatches_df[mismatches_df['Zone'] == zone]
                zone_df['End Time'] = zone_df['End Time'].replace("Not Closed", None)
                sorted_df = zone_df.sort_values(by='End Time', na_position='first')
                sorted_df['End Time'] = sorted_df['End Time'].fillna("Not Closed")

                message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                alias_col = 'Site Alias' if 'Site Alias' in sorted_df.columns else 'Site Alias '
                
                for site_alias in sorted_df[alias_col].dropna().unique():
                    site_df = sorted_df[sorted_df[alias_col] == site_alias]
                    message += f"✔ {site_alias}\n"
                    for _, row in site_df.iterrows():
                        message += f"  • Start: {row['Start Time']} | End: {row['End Time']}\n"
                    message += "\n"

                if zone in zone_to_name:
                    escaped_name = zone_to_name[zone].replace("_", "\\_")
                    message += f"**@{escaped_name}**, no Site Access Request found. Please take action.\n"

                requests.post(
                    f"https://api.telegram.org/bot{bot_token}/sendMessage",
                    json={"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
                )

# Streamlit App
st.title('🛡️IntrusionShield🛡️')

# File Uploaders
site_access_file = st.file_uploader("Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Historical Alarms Data", type=["xlsx"])
current_alarms_file = st.file_uploader("Current Alarms Data", type=["xlsx"])

# Initialize session state
if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    # Load data
    site_access_df = pd.read_excel(site_access_file)
    rms_df = pd.read_excel(rms_file, header=2)
    current_alarms_df = pd.read_excel(current_alarms_file, header=2)

    # Process alarms separately
    current_alarms_df['Start Time'] = current_alarms_df['Alarm Time']
    current_alarms_df['End Time'] = pd.NaT
    
    # Find mismatches for both types
    historical_mismatches = find_mismatches(site_access_df, rms_df)
    current_mismatches = find_mismatches(site_access_df, current_alarms_df)
    
    # Apply time filter
    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)
    historical_mismatches['Start Time'] = pd.to_datetime(historical_mismatches['Start Time'])
    current_mismatches['Start Time'] = pd.to_datetime(current_mismatches['Start Time'])
    
    filtered_historical = historical_mismatches[historical_mismatches['Start Time'] > filter_datetime]
    filtered_current = current_mismatches[current_mismatches['Start Time'] > filter_datetime]

    # Display both alarm types
    display_grouped_data(filtered_historical, "Historical Alarms")
    display_grouped_data(filtered_current, "Current Alarms")

    # Matched sites
    merged_alarms = merge_rms_alarms(rms_df, current_alarms_df)
    matched_df = find_matched_sites(site_access_df, merged_alarms)
    filtered_matched = matched_df[
        (matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True) &
        ((matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime))
    ]
    display_matched_sites(filtered_matched)

# Sidebar options
st.sidebar.title("Options")
user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

if os.path.exists(user_file_path):
    user_df = pd.read_excel(user_file_path)
    if "Zone" in user_df.columns and "Name" in user_df.columns:
        selected_zone = st.sidebar.selectbox("Zone", user_df['Zone'].unique())
        new_name = st.sidebar.text_input("Update Name", user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0])
        if st.sidebar.button("🔄 Update"):
            user_df.loc[user_df['Zone'] == selected_zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)

# Notification button
if st.sidebar.button("💬 Send Notification") and 'filtered_historical' in locals() and 'filtered_current' in locals():
    all_mismatches = pd.concat([filtered_historical, filtered_current])
    send_telegram_notification(all_mismatches, user_file_path)
