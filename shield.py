import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications
import os  # For file path operations
from io import BytesIO

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    if pd.isna(site_name):
        return site_name
    parts = str(site_name).split('_')
    return parts[0] if len(parts) > 1 else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    # Clean column names by stripping whitespace
    rms_df.columns = rms_df.columns.str.strip()
    alarms_df.columns = alarms_df.columns.str.strip()
    
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    # Select only the columns we need and ensure they exist
    rms_selected = rms_df[rms_columns] if all(col in rms_df.columns for col in rms_columns) else rms_df
    alarms_selected = alarms_df[alarms_columns] if all(col in alarms_df.columns for col in alarms_columns) else alarms_df
    
    merged_df = pd.concat([rms_selected, alarms_selected], ignore_index=True)
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(merged_df, site_access_df, 
                                  left_on='Site', 
                                  right_on='SiteName_Extracted', 
                                  how='left', 
                                  indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')  # Replace NaT with Not Closed
    return mismatches_df[['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']]

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    matched_df = pd.merge(site_access_df, merged_df, left_on='SiteName_Extracted', right_on='Site', how='inner')
    
    # Convert to datetime and handle errors
    matched_df['StartDate'] = pd.to_datetime(matched_df['StartDate'], errors='coerce')
    matched_df['EndDate'] = pd.to_datetime(matched_df['EndDate'], errors='coerce')
    matched_df['Start Time'] = pd.to_datetime(matched_df['Start Time'], errors='coerce')
    matched_df['End Time'] = pd.to_datetime(matched_df['End Time'], errors='coerce')
    
    # Determine status
    matched_df['Status'] = matched_df.apply(
        lambda row: 'Expired' if pd.notnull(row['End Time']) and pd.notnull(row['EndDate']) and (row['End Time'] > row['EndDate']) 
                    else 'Valid', 
        axis=1
    )
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
            
            # Ensure we have the correct columns and fill NA values
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            display_df = display_df.fillna('')
            
            st.table(display_df)
        st.markdown("---")

# Function to display matched sites with status
def display_matched_sites(matched_df):
    if not matched_df.empty:
        color_map = {'Valid': 'background-color: lightgreen;', 'Expired': 'background-color: lightcoral;'}
        
        def highlight_status(row):
            return [color_map.get(row['Status'], ''] * (len(row) - 1) + [color_map.get(row['Status'])]
        
        display_cols = ['RequestId', 'Site Alias', 'Start Time', 'End Time', 'EndDate', 'Status']
        display_cols = [col for col in display_cols if col in matched_df.columns]
        
        styled_df = matched_df[display_cols].style.apply(highlight_status, axis=1)
        st.write("Matched Sites with Status:")
        st.dataframe(styled_df)
    else:
        st.write("No matched sites found.")

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

# Function to update the user name for a specific zone
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)

        if "Zone" in user_df.columns and "Name" in user_df.columns:
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
        else:
            return False, "The USER NAME.xlsx file must have 'Zone' and 'Name' columns."
    else:
        return False, "USER NAME.xlsx file not found in the repository."

# Streamlit app
st.title('🛡️IntrusionShield🛡️')

# File uploaders
site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Door Open Alarms Data", type=["xlsx"])

# Initialize session state variables
if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    try:
        # Read data files
        site_access_df = pd.read_excel(site_access_file)
        rms_df = pd.read_excel(rms_file, header=2)
        current_alarms_df = pd.read_excel(current_alarms_file, header=2)
        
        # Clean column names by stripping whitespace
        rms_df.columns = rms_df.columns.str.strip()
        current_alarms_df.columns = current_alarms_df.columns.str.strip()
        site_access_df.columns = site_access_df.columns.str.strip()

        # Merge RMS and alarms data
        merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

        # Filter inputs (date and time)
        selected_date = st.date_input("Select Date", value=st.session_state.filter_date)
        selected_time = st.time_input("Select Time", value=st.session_state.filter_time)

        # Button to clear filters
        if st.button("Clear Filters"):
            st.session_state.filter_date = datetime.now().date()
            st.session_state.filter_time = datetime.now().time()
            st.session_state.status_filter = "All"
            selected_date = st.session_state.filter_date
            selected_time = st.session_state.filter_time

        # Update session state only when the user changes time or date
        if selected_date != st.session_state.filter_date:
            st.session_state.filter_date = selected_date
        if selected_time != st.session_state.filter_time:
            st.session_state.filter_time = selected_time

        # Combine selected date and time into a datetime object
        filter_datetime = datetime.combine(st.session_state.filter_date, st.session_state.filter_time)

        # Process mismatches
        mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
        mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
        filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

        # Process matches
        matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)

        # Status filter dropdown
        status_filter = st.selectbox("SA-Request Valid/Expired", options=["All", "Valid", "Expired"], index=0)
        
        # Update session state for status filter
        if status_filter != st.session_state.status_filter:
            st.session_state.status_filter = status_filter

        # Apply filtering conditions
        status_filter_condition = matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True
        time_filter_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)
        
        # Apply filters to matched data
        filtered_matched_df = matched_df[status_filter_condition & time_filter_condition]

        # Display mismatches
        if not filtered_mismatches_df.empty:
            st.write(f"Mismatched Sites (After {filter_datetime}) grouped by Cluster and Zone:")
            display_grouped_data(filtered_mismatches_df, "Filtered Mismatched Sites")
        else:
            st.write(f"No mismatches found after {filter_datetime}. Showing all mismatched sites.")
            display_grouped_data(mismatches_df, "All Mismatched Sites")

        # Display matched sites
        display_matched_sites(filtered_matched_df)

        # Sidebar options
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
        @st.cache_data
        def convert_df_to_excel_with_sheets(unmatched_df, rms_df, current_alarms_df, site_access_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                unmatched_df.to_excel(writer, index=False, sheet_name='Unmatched Data')
                rms_df.to_excel(writer, index=False, sheet_name='RMS Data')
                current_alarms_df.to_excel(writer, index=False, sheet_name='Current Alarms')
                site_access_df.to_excel(writer, index=False, sheet_name='Site Access Data')
            return output.getvalue()

        if not mismatches_df.empty:
            timestamp = datetime.now().strftime("%d%m%y%H%M%S")
            file_name = f"UnauthorizedAccess_{timestamp}.xlsx"
            excel_data = convert_df_to_excel_with_sheets(mismatches_df, rms_df, current_alarms_df, site_access_df)

            st.sidebar.download_button(
                label="📂 Download Data",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Telegram Notification Option
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
                        
                        # Clean data before creating message
                        zone_df['Site Alias'] = zone_df['Site Alias'].fillna('Unknown Site')
                        zone_df['Start Time'] = zone_df['Start Time'].dt.strftime('%Y-%m-%d %H:%M:%S')
                        zone_df['End Time'] = zone_df['End Time'].replace('Not Closed', 'Still Open')
                        
                        message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                        
                        for _, row in zone_df.iterrows():
                            message += f"✔ {row['Site Alias']}\n"
                            message += f"  • Start Time: {row['Start Time']} | End Time: {row['End Time']}\n\n"
                        
                        if zone in zone_to_name:
                            escaped_name = str(zone_to_name[zone]).replace("_", "\\_")
                            message += f"**@{escaped_name}**, no Site Access Request found for these Door Open alarms. Please take care and share us update.\n"
                        
                        if send_telegram_notification(message, bot_token, chat_id):
                            st.success(f"Notification for zone '{zone}' sent successfully!")
                        else:
                            st.error(f"Failed to send notification for zone '{zone}'.")
                else:
                    st.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
            else:
                st.error("USER NAME.xlsx file not found in the repository.")

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
