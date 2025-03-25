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
    # Standardize column names
    rms_df = rms_df.rename(columns={'Site Alias ': 'Site Alias'})
    alarms_df = alarms_df.rename(columns={'Alarm Time': 'Start Time'})
    
    alarms_df['End Time'] = pd.NaT
    return pd.concat([
        rms_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']],
        alarms_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]
    ], ignore_index=True)

def find_mismatches(site_access_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged = pd.merge(alarms_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches = merged[merged['_merge'] == 'left_only']
    mismatches['End Time'] = mismatches['End Time'].fillna('Not Closed')
    return mismatches

def display_grouped_data(grouped_df, title):
    st.write(title)
    
    # Handle empty dataframe case
    if grouped_df.empty:
        st.write("No mismatches found")
        return
        
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            display_df = display_df.fillna('Unknown')
            
            # Only show Site Alias when it changes
            display_df['Site Alias'] = display_df['Site Alias'].where(~display_df['Site Alias'].duplicated(), '')
            
            st.table(display_df.reset_index(drop=True))
        st.markdown("---")

def send_telegram_notification(mismatches_df, user_file_path):
    if not os.path.exists(user_file_path):
        st.error("USER NAME.xlsx file not found in the repository.")
        return

    user_df = pd.read_excel(user_file_path)
    if "Zone" not in user_df.columns or "Name" not in user_df.columns:
        st.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
        return

    zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
    bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
    chat_id = "-4625672098"

    for zone in mismatches_df['Zone'].unique():
        zone_df = mismatches_df[mismatches_df['Zone'] == zone]
        
        message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
        
        for _, row in zone_df.iterrows():
            site_alias = row['Site Alias'] if pd.notna(row['Site Alias']) else "Unknown"
            start_time = row['Start Time']
            end_time = row['End Time'] if pd.notna(row['End Time']) else "Not Closed"
            
            message += f"✔ {site_alias}\n"
            message += f"  • Start Time: {start_time} | End Time: {end_time}\n\n"

        if zone in zone_to_name:
            escaped_name = zone_to_name[zone].replace("_", "\\_")
            message += f"**@{escaped_name}**, no Site Access Request found for these Door Open alarms. Please take care and share update.\n"

        response = requests.post(
            f"https://api.telegram.org/bot{bot_token}/sendMessage",
            json={"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
        )
        
        if response.status_code == 200:
            st.success(f"Notification for zone '{zone}' sent successfully!")
        else:
            st.error(f"Failed to send notification for zone '{zone}'.")

# Streamlit App
st.title('🛡️IntrusionShield🛡️')

# File Uploaders
site_access_file = st.file_uploader("Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Current Door Open Alarms Data", type=["xlsx"])

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

    # Process and merge all alarms
    merged_alarms = merge_rms_alarms(rms_df, current_alarms_df)
    
    # Find all mismatches
    all_mismatches = find_mismatches(site_access_df, merged_alarms)
    
    # Apply time filter
    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)
    all_mismatches['Start Time'] = pd.to_datetime(all_mismatches['Start Time'])
    filtered_mismatches = all_mismatches[all_mismatches['Start Time'] > filter_datetime]

    # Display all mismatches together
    if not filtered_mismatches.empty:
        display_grouped_data(filtered_mismatches, "Mismatched Sites")
    else:
        display_grouped_data(all_mismatches, "All Mismatched Sites")

    # Sidebar notification button
    if st.sidebar.button("💬 Send Notification"):
        send_telegram_notification(filtered_mismatches if not filtered_mismatches.empty else all_mismatches, 
                                 os.path.join(os.path.dirname(__file__), "USER NAME.xlsx"))
