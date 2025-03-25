import pandas as pd
import streamlit as st
from datetime import datetime
import requests
import os
from io import BytesIO

# Function to extract site name
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms
def merge_rms_alarms(rms_df, alarms_df):
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT
    return pd.concat([
        rms_df[['Site', 'Site Alias ', 'Zone', 'Cluster', 'Start Time', 'End Time']],
        alarms_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]
    ], ignore_index=True)

# Function to find mismatches
def find_mismatches(site_access_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged = pd.merge(alarms_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches = merged[merged['_merge'] == 'left_only']
    mismatches['End Time'] = mismatches['End Time'].fillna('Not Closed')
    return mismatches

# Function to find matched sites
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched['StartDate'] = pd.to_datetime(matched['StartDate'], errors='coerce')
    matched['EndDate'] = pd.to_datetime(matched['EndDate'], errors='coerce')
    matched['Start Time'] = pd.to_datetime(matched['Start Time'], errors='coerce')
    matched['End Time'] = pd.to_datetime(matched['End Time'], errors='coerce')
    matched['Status'] = matched.apply(lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched

# Fixed display function
def display_grouped_data(grouped_df, title, alarm_type=""):
    st.write(f"### {title} {alarm_type}")
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            # Handle column name variations
            alias_col = 'Site Alias' if 'Site Alias' in zone_df.columns else 'Site Alias '
            display_df = zone_df[[alias_col, 'Start Time', 'End Time']].copy()
            display_df.columns = ['Site Alias', 'Start Time', 'End Time']
            
            # Clean and format data
            display_df['Site Alias'] = display_df['Site Alias'].fillna('Unknown').replace('', 'Unknown')
            display_df = display_df.reset_index(drop=True)
            
            # Only show Site Alias when it changes
            prev_alias = None
            for i in range(len(display_df)):
                current_alias = display_df.at[i, 'Site Alias']
                if current_alias == prev_alias:
                    display_df.at[i, 'Site Alias'] = ''
                else:
                    prev_alias = current_alias
            
            st.table(display_df)
        st.markdown("---")

# Display matched sites
def display_matched_sites(matched_df):
    color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
    def highlight_status(status):
        return color_map.get(status, '')
    styled_df = matched_df[['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']].style.applymap(highlight_status, subset=['Status'])
    st.write("Matched Sites with Status:")
    st.dataframe(styled_df)

# Fixed Telegram notification
def send_telegram_notification(historical_df, current_df, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
            bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
            chat_id = "-4625672098"

            # Combine both dataframes
            all_alarms = pd.concat([historical_df, current_df])
            
            for zone in all_alarms['Zone'].unique():
                zone_df = all_alarms[all_alarms['Zone'] == zone]
                message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                
                # Handle column name variations
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
                    message += f"**@{escaped_name}**, no Site Access Request found.\n"

                response = requests.post(
                    f"https://api.telegram.org/bot{bot_token}/sendMessage",
                    json={"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}
                )

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

    # Process historical alarms
    rms_mismatches = find_mismatches(site_access_df, rms_df)
    rms_mismatches['Start Time'] = pd.to_datetime(rms_mismatches['Start Time'])
    
    # Process current alarms
    current_alarms_df['Start Time'] = current_alarms_df['Alarm Time']
    current_mismatches = find_mismatches(site_access_df, current_alarms_df)
    current_mismatches['Start Time'] = pd.to_datetime(current_mismatches['Start Time'])
    
    # Apply time filter
    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)
    filtered_rms = rms_mismatches[rms_mismatches['Start Time'] > filter_datetime]
    filtered_current = current_mismatches[current_mismatches['Start Time'] > filter_datetime]

    # Display both alarm types separately
    display_grouped_data(filtered_rms if not filtered_rms.empty else rms_mismatches, 
                        "Historical Alarms", "(From All Door Open Alarms Data)")
    display_grouped_data(filtered_current if not filtered_current.empty else current_mismatches, 
                        "Current Alarms", "(From Current Door Open Alarms Data)")

    # Matched sites display
    merged_alarms = merge_rms_alarms(rms_df, current_alarms_df)
    matched_df = find_matched_sites(site_access_df, merged_alarms)
    status_filter = st.selectbox("SA-Request Valid/Expired", options=["All", "Valid", "Expired"])
    filtered_matched = matched_df[
        (matched_df['Status'] == status_filter if status_filter != "All" else True) &
        ((matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime))
    ]
    display_matched_sites(filtered_matched)

# [Rest of your original code (download, zone update etc.) remains exactly the same]

# [Previous code remains exactly the same until the matched sites display]

    # Display matched sites with status
    display_matched_sites(filtered_matched_df)

# Function to update zone user
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
        return False, "The USER NAME.xlsx file must have 'Zone' and 'Name' columns."
    return False, "USER NAME.xlsx file not found in the repository."

# Download function
@st.cache_data
def convert_df_to_excel_with_sheets(historical_df, current_df, rms_df, current_alarms_df, site_access_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Historical alarms sheet
        historical_df[['Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']].to_excel(
            writer, index=False, sheet_name='Historical Alarms')
        
        # Current alarms sheet
        current_df[['Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']].to_excel(
            writer, index=False, sheet_name='Current Alarms')
        
        # Raw data sheets
        rms_df.to_excel(writer, index=False, sheet_name='RMS Raw Data')
        current_alarms_df.to_excel(writer, index=False, sheet_name='Current Alarms Raw Data')
        site_access_df.to_excel(writer, index=False, sheet_name='Site Access Data')

        # Formatting
        workbook = writer.book
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(writer.sheets[sheet_name]._cols):
                max_len = max(len(str(col.header)), *[len(str(x)) for x in col.col_values])
                worksheet.set_column(i, i, max_len + 2)
    return output.getvalue()

# Sidebar options
st.sidebar.title("Options")
user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

if os.path.exists(user_file_path):
    user_df = pd.read_excel(user_file_path)
    if "Zone" in user_df.columns and "Name" in user_df.columns:
        selected_zone = st.sidebar.selectbox("Select Zone", user_df['Zone'].unique())
        current_name = user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0]
        new_name = st.sidebar.text_input("Update Name", value=current_name)
        if st.sidebar.button("🔄 Update Concern"):
            success, message = update_zone_user(selected_zone, new_name, user_file_path)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.error(message)
    else:
        st.sidebar.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
else:
    st.sidebar.error("USER NAME.xlsx file not found in the repository.")

# Download button
if site_access_file and rms_file and current_alarms_file:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_data = convert_df_to_excel_with_sheets(
        filtered_rms if not filtered_rms.empty else rms_mismatches,
        filtered_current if not filtered_current.empty else current_mismatches,
        rms_df,
        current_alarms_df,
        site_access_df
    )
    st.sidebar.download_button(
        label="📥 Download Full Report",
        data=excel_data,
        file_name=f"IntrusionShield_Report_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Telegram notification button
if st.sidebar.button("💬 Send Notification"):
    send_telegram_notification(
        filtered_rms if not filtered_rms.empty else rms_mismatches,
        filtered_current if not filtered_current.empty else current_mismatches,
        user_file_path
    )
