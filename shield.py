import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications
import os  # For file path operations

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to process RMS and Current Alarms data separately
def process_alarms_data(rms_df, alarms_df, site_access_df):
    # Process historical alarms (RMS data)
    rms_df['Start Time'] = pd.to_datetime(rms_df['Start Time'])
    rms_df['End Time'] = pd.to_datetime(rms_df['End Time'])
    historical_mismatches = find_mismatches(site_access_df, rms_df)
    
    # Process current alarms
    alarms_df['Start Time'] = pd.to_datetime(alarms_df['Alarm Time'])
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms
    current_mismatches = find_mismatches(site_access_df, alarms_df)
    
    return historical_mismatches, current_mismatches

# Function to find mismatches between Site Access and alarms dataset
def find_mismatches(site_access_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(alarms_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')  # Replace NaT with Not Closed
    return mismatches_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, rms_df, alarms_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    
    # Combine both alarm types for matching
    combined_alarms = pd.concat([
        rms_df[['Site', 'Site Alias ', 'Zone', 'Cluster', 'Start Time', 'End Time']].rename(columns={'Site Alias ': 'Site Alias'}),
        alarms_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Alarm Time', 'End Time']].rename(columns={'Alarm Time': 'Start Time'})
    ])
    
    matched_df = pd.merge(site_access_df, combined_alarms, left_on='SiteName_Extracted', right_on='Site', how='inner')
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')
    matched_df['Status'] = matched_df.apply(lambda row: 'Expired' if pd.notnull(row['End Time']) and row['End Time'] > row['EndDate'] else 'Valid', axis=1)
    return matched_df

# Function to display grouped data by Cluster and Zone in a table
def display_grouped_data(grouped_df, title, alarm_type):
    st.write(f"### {title} ({alarm_type})")
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[grouped_df['Zone'] == zone]
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
def send_telegram_notification(historical_mismatches, current_mismatches, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)

        if "Zone" in user_df.columns and "Name" in user_df.columns:
            zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
            bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
            chat_id = "-4625672098"

            # Process historical alarms
            if not historical_mismatches.empty:
                historical_zones = historical_mismatches['Zone'].unique()
                for zone in historical_zones:
                    zone_df = historical_mismatches[historical_mismatches['Zone'] == zone]
                    message = f"❗Historical Door Open Notification❗\n\n🚩 {zone}\n\n"
                    
                    site_aliases = zone_df['Site Alias'].unique()
                    for site_alias in site_aliases:
                        site_df = zone_df[zone_df['Site Alias'] == site_alias]
                        message += f"✔ {site_alias}\n"
                        for _, row in site_df.iterrows():
                            end_time_display = row['End Time']
                            message += f"  • Start Time: {row['Start Time']} | End Time: {end_time_display}\n"
                        message += "\n"

                    if zone in zone_to_name:
                        escaped_name = zone_to_name[zone].replace("_", "\\_")
                        message += f"**@{escaped_name}**, no Site Access Request found for these historical Door Open alarms.\n"

                    payload = {
                        "chat_id": chat_id,
                        "text": message,
                        "parse_mode": "Markdown"
                    }
                    response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendMessage", json=payload)
                    if response.status_code == 200:
                        st.success(f"Historical notification for zone '{zone}' sent successfully!")
                    else:
                        st.error(f"Failed to send historical notification for zone '{zone}'.")

            # Process current alarms
            if not current_mismatches.empty:
                current_zones = current_mismatches['Zone'].unique()
                for zone in current_zones:
                    zone_df = current_mismatches[current_mismatches['Zone'] == zone]
                    message = f"❗Current Door Open Notification❗\n\n🚩 {zone}\n\n"
                    
                    site_aliases = zone_df['Site Alias'].unique()
                    for site_alias in site_aliases:
                        site_df = zone_df[zone_df['Site Alias'] == site_alias]
                        message += f"✔ {site_alias}\n"
                        for _, row in site_df.iterrows():
                            end_time_display = row['End Time']
                            message += f"  • Start Time: {row['Start Time']} | End Time: {end_time_display}\n"
                        message += "\n"

                    if zone in zone_to_name:
                        escaped_name = zone_to_name[zone].replace("_", "\\_")
                        message += f"**@{escaped_name}**, no Site Access Request found for these CURRENT Door Open alarms. Please take immediate action!\n"

                    payload = {
                        "chat_id": chat_id,
                        "text": message,
                        "parse_mode": "Markdown"
                    }
                    response = requests.post(f"https://api.telegram.org/bot{bot_token}/sendMessage", json=payload)
                    if response.status_code == 200:
                        st.success(f"Current notification for zone '{zone}' sent successfully!")
                    else:
                        st.error(f"Failed to send current notification for zone '{zone}'.")
        else:
            st.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
    else:
        st.error("USER NAME.xlsx file not found in the repository.")

# Streamlit app
st.title('🛡️IntrusionShield🛡️')

site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now (Historical)", type=["xlsx"])
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

    # Process alarms data separately
    historical_mismatches, current_mismatches = process_alarms_data(rms_df, current_alarms_df, site_access_df)
    
    # Find matched sites
    matched_df = find_matched_sites(site_access_df, rms_df, current_alarms_df)

    # Filter inputs (date and time)
    selected_date = st.date_input("Select Date", value=st.session_state.filter_date)
    selected_time = st.time_input("Select Time", value=st.session_state.filter_time)

    # Button to clear filters
    if st.button("Clear Filters"):
        st.session_state.filter_date = datetime.now().date()
        st.session_state.filter_time = datetime.now().time()
        st.session_state.status_filter = "All"

    # Update session state only when the user changes time or date
    if selected_date != st.session_state.filter_date:
        st.session_state.filter_date = selected_date
    if selected_time != st.session_state.filter_time:
        st.session_state.filter_time = selected_time

    # Combine selected date and time into a datetime object
    filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)

    # Apply time filter to both historical and current alarms
    historical_mismatches['Start Time'] = pd.to_datetime(historical_mismatches['Start Time'], errors='coerce')
    filtered_historical = historical_mismatches[historical_mismatches['Start Time'] > filter_datetime]
    
    current_mismatches['Start Time'] = pd.to_datetime(current_mismatches['Start Time'], errors='coerce')
    filtered_current = current_mismatches[current_mismatches['Start Time'] > filter_datetime]

    # Apply filters to matched data
    status_filter_condition = matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True
    time_filter_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)
    filtered_matched_df = matched_df[status_filter_condition & time_filter_condition]

    # Add the status filter dropdown right before the matched sites table
    status_filter = st.selectbox("SA-Request Valid/Expired", options=["All", "Valid", "Expired"], index=0)

    # Update session state for status filter
    if status_filter != st.session_state.status_filter:
        st.session_state.status_filter = status_filter

    # Display historical mismatches
    if not filtered_historical.empty:
        display_grouped_data(filtered_historical, "Historical Alarms", "Historical")
    else:
        st.write(f"No historical mismatches found after {filter_datetime}. Showing all historical mismatched sites.")
        display_grouped_data(historical_mismatches, "All Historical Alarms", "Historical")

    # Display current mismatches
    if not filtered_current.empty:
        display_grouped_data(filtered_current, "Current Alarms", "Current")
    else:
        st.write(f"No current mismatches found after {filter_datetime}. Showing all current mismatched sites.")
        display_grouped_data(current_mismatches, "All Current Alarms", "Current")

    # Display matched sites
    display_matched_sites(filtered_matched_df)

# Function to update the user name for a specific zone
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)

        # Ensure proper column names
        if "Zone" in user_df.columns and "Name" in user_df.columns:
            # Update the name for the selected zone
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name

            # Save the updated DataFrame back to the file
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
        else:
            return False, "The USER NAME.xlsx file must have 'Zone' and 'Name' columns."
    else:
        return False, "USER NAME.xlsx file not found in the repository."

# Streamlit Sidebar
st.sidebar.title("Options")

# Update Zone Concern Option
st.sidebar.markdown("### Update Zone Concern")
user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

if os.path.exists(user_file_path):
    user_df = pd.read_excel(user_file_path)

    if "Zone" in user_df.columns and "Name" in user_df.columns:
        zone_list = user_df['Zone'].unique()
        selected_zone = st.sidebar.selectbox("Select Zone", options=zone_list)

        if selected_zone:
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
from io import BytesIO
from datetime import datetime

@st.cache_data
def convert_df_to_excel_with_sheets(historical_mismatches, current_mismatches, rms_df, current_alarms_df, site_access_df):
    # Filter data to show only the required columns
    filtered_historical = historical_mismatches[['Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]
    filtered_current = current_mismatches[['Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]

    # Create an Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Add historical alarms sheet
        filtered_historical.to_excel(writer, index=False, sheet_name='Historical Alarms')
        
        # Add current alarms sheet
        filtered_current.to_excel(writer, index=False, sheet_name='Current Alarms')

        # Add raw RMS Data sheet
        rms_df.to_excel(writer, index=False, sheet_name='RMS Data')

        # Add raw Current Alarms sheet
        current_alarms_df.to_excel(writer, index=False, sheet_name='Raw Current Alarms')

        # Add raw Site Access Data sheet
        site_access_df.to_excel(writer, index=False, sheet_name='Site Access Data')

        # Access the workbook for formatting
        workbook = writer.book

        # Format each sheet with auto-adjusted column widths and table style
        for sheet_name, df in [
            ('Historical Alarms', filtered_historical),
            ('Current Alarms', filtered_current),
            ('RMS Data', rms_df),
            ('Raw Current Alarms', current_alarms_df),
            ('Site Access Data', site_access_df)
        ]:
            worksheet = writer.sheets[sheet_name]
            for i, column in enumerate(df.columns):
                max_len = max(df[column].astype(str).map(len).max(), len(column)) + 2
                worksheet.set_column(i, i, max_len)

            # Apply table formatting to alarms sheets
            if sheet_name in ['Historical Alarms', 'Current Alarms']:
                table_range = f'A1:E{len(df) + 1}'  # Adjust range for headers and data
                worksheet.add_table(table_range, {
                    'columns': [{'header': col} for col in df.columns],
                    'style': 'Table Style Medium 9',
                })

    return output.getvalue()

# Generate the Excel file only if there is data
if site_access_file and rms_file and current_alarms_file:
    # Generate the file name with current timestamp
    timestamp = datetime.now().strftime("%d%m%y%H%M%S")
    file_name = f"UnauthorizedAccess_{timestamp}.xlsx"

    # Generate the Excel data with all sheets
    excel_data = convert_df_to_excel_with_sheets(historical_mismatches, current_mismatches, rms_df, current_alarms_df, site_access_df)

    # Add a download button in the sidebar
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
    send_telegram_notification(historical_mismatches, current_mismatches, user_file_path)
