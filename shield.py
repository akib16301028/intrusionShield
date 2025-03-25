import pandas as pd
import streamlit as st
from datetime import datetime
import requests
import os
from io import BytesIO

# Function to extract site name
def extract_site(site_name):
    if pd.isna(site_name):
        return site_name
    if isinstance(site_name, str) and '_' in site_name:
        return site_name.split('_')[0]
    return site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    # Standardize column names
    rms_df.columns = rms_df.columns.str.strip()
    alarms_df.columns = alarms_df.columns.str.strip()
    
    # Handle column name variations
    site_alias_col = 'Site Alias' if 'Site Alias' in alarms_df.columns else 'Site Alias '
    
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT

    rms_columns = ['Site', site_alias_col, 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', site_alias_col, 'Zone', 'Cluster', 'Start Time', 'End Time']

    merged_df = pd.concat([
        rms_df[rms_columns].rename(columns={site_alias_col: 'Site Alias'}),
        alarms_df[alarms_columns].rename(columns={site_alias_col: 'Site Alias'})
    ], ignore_index=True)
    
    return merged_df

# Function to find mismatches
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(
        merged_df, 
        site_access_df, 
        left_on='Site', 
        right_on='SiteName_Extracted', 
        how='left', 
        indicator=True
    )
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')
    
    return mismatches_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]

# Function to find matched sites
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(
        site_access_df, 
        merged_df, 
        left_on='SiteName_Extracted', 
        right_on='Site', 
        how='inner'
    )
    
    # Convert to datetime
    for col in ['StartDate', 'EndDate', 'Start Time', 'End Time']:
        matched_df[col] = pd.to_datetime(matched_df[col], errors='coerce')
    
    # Determine status
    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if (pd.notnull(row['End Time']) and 
                                pd.notnull(row['EndDate']) and 
                                (row['End Time'] > row['EndDate'])) 
                   else 'Valid', 
        axis=1
    )
    return matched_df

# Function to display grouped data
def display_grouped_data(grouped_df, title):
    st.write(title)
    
    if 'Site Alias' not in grouped_df.columns:
        st.error("Site Alias column not found in data")
        return
    
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***{zone}***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            display_df = display_df.fillna('')
            st.table(display_df)
        st.markdown("---")

# Function to display matched sites with status
def display_matched_sites(matched_df):
    if not matched_df.empty:
        def highlight_status(row):
            color = 'lightgreen' if row['Status'] == 'Valid' else 'lightcoral'
            return [f'background-color: {color}'] * len(row)
        
        display_cols = ['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']
        display_cols = [col for col in display_cols if col in matched_df.columns]
        
        styled_df = matched_df[display_cols].style.apply(highlight_status, axis=1)
        st.dataframe(styled_df)
    else:
        st.write("No matched sites found.")

# Function to update zone user
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
    return False, "Error updating zone concern"

# Function to send Telegram notification
def send_telegram_notification(message, bot_token, chat_id):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "Markdown"
    }
    try:
        response = requests.post(url, json=payload)
        return response.status_code == 200
    except Exception as e:
        st.error(f"Error sending notification: {e}")
        return False

# Streamlit UI
st.title('🛡️IntrusionShield🛡️')

# File uploaders
site_access_file = st.file_uploader("Upload Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload All Door Open Alarms Data", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload Current Door Open Alarms Data", type=["xlsx"])

# Initialize session state
if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    try:
        # Read and clean data
        site_access_df = pd.read_excel(site_access_file)
        rms_df = pd.read_excel(rms_file, header=2)
        current_alarms_df = pd.read_excel(current_alarms_file, header=2)
        
        # Standardize column names
        for df in [site_access_df, rms_df, current_alarms_df]:
            df.columns = df.columns.str.strip()

        # Merge data
        merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

        # Date/time filters
        selected_date = st.date_input("Select Date", value=st.session_state.filter_date)
        selected_time = st.time_input("Select Time", value=st.session_state.filter_time)
        
        if st.button("Clear Filters"):
            st.session_state.filter_date = datetime.now().date()
            st.session_state.filter_time = datetime.now().time()
            st.session_state.status_filter = "All"
            selected_date = st.session_state.filter_date
            selected_time = st.session_state.filter_time

        filter_datetime = datetime.combine(selected_date, selected_time)

        # Process data
        mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
        mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
        filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

        matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)
        
        # Status filter
        status_filter = st.selectbox(
            "SA-Request Valid/Expired", 
            options=["All", "Valid", "Expired"], 
            index=["All", "Valid", "Expired"].index(st.session_state.status_filter)
        )
        st.session_state.status_filter = status_filter

        # Apply filters
        status_condition = (matched_df['Status'] == status_filter) if status_filter != "All" else True
        time_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)
        filtered_matched_df = matched_df[status_condition & time_condition]

        # Display results
        if not filtered_mismatches_df.empty:
            display_grouped_data(filtered_mismatches_df, f"Mismatched Sites (After {filter_datetime})")
        else:
            display_grouped_data(mismatches_df, "All Mismatched Sites")

        display_matched_sites(filtered_matched_df)

        # Sidebar options
        st.sidebar.title("Options")
        
        # Update Zone Concern
        st.sidebar.markdown("### Update Zone Concern")
        user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")
        
        if os.path.exists(user_file_path):
            user_df = pd.read_excel(user_file_path)
            if "Zone" in user_df.columns and "Name" in user_df.columns:
                zone_list = user_df['Zone'].unique()
                selected_zone = st.sidebar.selectbox("Select Zone", zone_list)
                
                if selected_zone:
                    current_name = user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0]
                    new_name = st.sidebar.text_input("Update Name", value=current_name)
                    
                    if st.sidebar.button("🔄 Update Concern"):
                        success, message = update_zone_user(selected_zone, new_name, user_file_path)
                        if success:
                            st.sidebar.success(message)
                        else:
                            st.sidebar.error(message)

        # Download Option
        @st.cache_data
        def convert_to_excel(unmatched_df, rms_df, current_alarms_df, site_access_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                unmatched_df.to_excel(writer, sheet_name='Unmatched Data', index=False)
                rms_df.to_excel(writer, sheet_name='RMS Data', index=False)
                current_alarms_df.to_excel(writer, sheet_name='Current Alarms', index=False)
                site_access_df.to_excel(writer, sheet_name='Site Access Data', index=False)
            return output.getvalue()

        if not mismatches_df.empty:
            excel_data = convert_to_excel(mismatches_df, rms_df, current_alarms_df, site_access_df)
            st.sidebar.download_button(
                label="📂 Download Data",
                data=excel_data,
                file_name=f"UnauthorizedAccess_{datetime.now().strftime('%d%m%y%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Telegram Notification
        if st.sidebar.button("💬 Send Notification"):
            if os.path.exists(user_file_path):
                user_df = pd.read_excel(user_file_path)
                if "Zone" in user_df.columns and "Name" in user_df.columns:
                    zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
                    bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
                    chat_id = "-4625672098"

                    zones = filtered_mismatches_df['Zone'].unique()
                    
                    for zone in zones:
                        zone_df = filtered_mismatches_df[filtered_mismatches_df['Zone'] == zone]
                        
                        # Prepare message
                        message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                        
                        for _, row in zone_df.iterrows():
                            site_alias = row.get('Site Alias', 'Unknown Site')
                            start_time = row.get('Start Time', 'Unknown time')
                            end_time = row.get('End Time', 'Still open')
                            
                            if isinstance(start_time, pd.Timestamp):
                                start_time = start_time.strftime('%Y-%m-%d %H:%M:%S')
                            
                            message += f"✔ {site_alias}\n"
                            message += f"  • Start Time: {start_time} | End Time: {end_time}\n\n"
                        
                        if zone in zone_to_name:
                            name = str(zone_to_name[zone]).replace("_", "\\_")
                            message += f"**@{name}**, no Site Access Request found. Please check and update.\n"
                        
                        if send_telegram_notification(message, bot_token, chat_id):
                            st.success(f"Notification for {zone} sent!")
                        else:
                            st.error(f"Failed to send notification for {zone}")
                
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
