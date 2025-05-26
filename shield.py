import pandas as pd
import streamlit as st
from datetime import datetime
import requests  # For sending Telegram notifications
import os  # For file path operations
from io import BytesIO

# Function to standardize column names by stripping whitespace
def clean_column_names(df):
    df.columns = df.columns.str.strip()
    return df

# Function to extract the first part of the SiteName before the first underscore
def extract_site(site_name):
    return site_name.split('_')[0] if pd.notnull(site_name) and '_' in site_name else site_name

# Function to merge RMS and Current Alarms data
def merge_rms_alarms(rms_df, alarms_df):
    # Clean column names
    rms_df = clean_column_names(rms_df)
    alarms_df = clean_column_names(alarms_df)
    
    alarms_df['Start Time'] = alarms_df['Alarm Time']
    alarms_df['End Time'] = pd.NaT  # No End Time in Current Alarms, set to NaT

    rms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']
    alarms_columns = ['Site', 'Site Alias', 'Zone', 'Cluster', 'Start Time', 'End Time']

    merged_df = pd.concat([rms_df[rms_columns], alarms_df[alarms_columns]], ignore_index=True)
    return merged_df

# Function to find mismatches between Site Access and merged RMS/Alarms dataset
def find_mismatches(site_access_df, merged_df):
    # Clean column names
    site_access_df = clean_column_names(site_access_df)
    merged_df = clean_column_names(merged_df)
    
    site_access_df['SiteName_Extracted'] = site_access_df['SiteName'].apply(extract_site)
    merged_comparison_df = pd.merge(merged_df, site_access_df, left_on='Site', right_on='SiteName_Extracted', how='left', indicator=True)
    mismatches_df = merged_comparison_df[merged_comparison_df['_merge'] == 'left_only']
    mismatches_df['End Time'] = mismatches_df['End Time'].fillna('Not Closed')  # Replace NaT with Not Closed
    return mismatches_df

# Function to find matched sites and their status
def find_matched_sites(site_access_df, merged_df):
    # Clean column names
    site_access_df = clean_column_names(site_access_df)
    merged_df = clean_column_names(merged_df)
    
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
    
    # Clean column names
    grouped_df = clean_column_names(grouped_df)
    
    clusters = grouped_df['Cluster'].unique()

    for cluster in clusters:
        st.markdown(f"**{cluster}**")
        cluster_df = grouped_df[grouped_df['Cluster'] == cluster]
        zones = cluster_df['Zone'].unique()

        for zone in zones:
            st.markdown(f"***<span style='font-size:14px;'>{zone}</span>***", unsafe_allow_html=True)
            zone_df = cluster_df[cluster_df['Zone'] == zone]
            
            # Create display DataFrame with cleaned column names
            display_df = zone_df[['Site Alias', 'Start Time', 'End Time']].copy()
            
            # Reset index for clean display
            display_df = display_df.reset_index(drop=True)
            
            # Only show Site Alias when it changes from the previous row OR when End Time is "Not Closed"
            for i in range(1, len(display_df)):
                if display_df.at[i, 'Site Alias'] == display_df.at[i-1, 'Site Alias'] and display_df.at[i, 'End Time'] != "Not Closed":
                    display_df.at[i, 'Site Alias'] = ''
            
            st.table(display_df)
        st.markdown("---")

# Function to display matched sites with status
def display_matched_sites(matched_df):
    # Clean column names
    matched_df = clean_column_names(matched_df)
    
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
        "parse_mode": "Markdown"
    }
    response = requests.post(url, json=payload)
    return response.status_code == 200

# Function to update the user name for a specific zone
def update_zone_user(zone, new_name, user_file_path):
    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        user_df = clean_column_names(user_df)
        user_df['Name'] = user_df['Name'].astype(str)  # Ensure names are strings

        if "Zone" in user_df.columns and "Name" in user_df.columns:
            user_df.loc[user_df['Zone'] == zone, 'Name'] = new_name
            user_df.to_excel(user_file_path, index=False)
            return True, "Zone concern updated successfully!"
        else:
            return False, "The USER NAME.xlsx file must have 'Zone' and 'Name' columns."
    else:
        return False, "USER NAME.xlsx file not found in the repository."

# Function to convert dataframes into an Excel file with multiple sheets
@st.cache_data
def convert_df_to_excel_with_sheets(unmatched_df, rms_df, current_alarms_df, site_access_df):
    # Clean column names
    unmatched_df = clean_column_names(unmatched_df)
    rms_df = clean_column_names(rms_df)
    current_alarms_df = clean_column_names(current_alarms_df)
    site_access_df = clean_column_names(site_access_df)
    
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

# File uploaders
site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Door Open Alarms Data", type=["xlsx"])

# Initialize session state
if "filter_time" not in st.session_state:
    st.session_state.filter_time = datetime.now().time()
if "filter_date" not in st.session_state:
    st.session_state.filter_date = datetime.now().date()
if "status_filter" not in st.session_state:
    st.session_state.status_filter = "All"

if site_access_file and rms_file and current_alarms_file:
    # Read and clean data
    site_access_df = clean_column_names(pd.read_excel(site_access_file))
    rms_df = clean_column_names(pd.read_excel(rms_file, header=2))
    current_alarms_df = clean_column_names(pd.read_excel(current_alarms_file, header=2))

    merged_rms_alarms_df = merge_rms_alarms(rms_df, current_alarms_df)

    # Filter inputs
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

    # Process data
    mismatches_df = find_mismatches(site_access_df, merged_rms_alarms_df)
    mismatches_df['Start Time'] = pd.to_datetime(mismatches_df['Start Time'], errors='coerce')
    filtered_mismatches_df = mismatches_df[mismatches_df['Start Time'] > filter_datetime]

    matched_df = find_matched_sites(site_access_df, merged_rms_alarms_df)
    status_filter_condition = matched_df['Status'] == st.session_state.status_filter if st.session_state.status_filter != "All" else True
    time_filter_condition = (matched_df['Start Time'] > filter_datetime) | (matched_df['End Time'] > filter_datetime)
    filtered_matched_df = matched_df[status_filter_condition & time_filter_condition]

    # Status filter
    status_filter = st.selectbox("SA-Request Valid/Expired", options=["All", "Valid", "Expired"], index=0)
    if status_filter != st.session_state.status_filter:
        st.session_state.status_filter = status_filter

    # Display results
    if not filtered_mismatches_df.empty:
        st.write(f"Mismatched Sites (After {filter_datetime}) grouped by Cluster and Zone:")
        display_grouped_data(filtered_mismatches_df, "Filtered Mismatched Sites")
    else:
        st.write(f"No mismatches found after {filter_datetime}. Showing all mismatched sites.")
        display_grouped_data(mismatches_df, "All Mismatched Sites")

    display_matched_sites(filtered_matched_df)

    # Sidebar operations
    st.sidebar.title("Options")
    st.sidebar.markdown("### Update Zone Concern")
    user_file_path = os.path.join(os.path.dirname(__file__), "USER NAME.xlsx")

    if os.path.exists(user_file_path):
        user_df = pd.read_excel(user_file_path)
        user_df = clean_column_names(user_df)
        user_df['Name'] = user_df['Name'].astype(str)  # Ensure names are strings

        if "Zone" in user_df.columns and "Name" in user_df.columns:
            zone_list = user_df['Zone'].unique()
            selected_zone = st.sidebar.selectbox("Select Zone", options=zone_list)

            if selected_zone:
                current_name = user_df.loc[user_df['Zone'] == selected_zone, 'Name'].values[0]
                new_name = st.sidebar.text_input("Update Name", value=current_name)

                if st.sidebar.button("🔄Update Concern"):
                    success, message = update_zone_user(selected_zone, str(new_name), user_file_path)
                    if success:
                        st.sidebar.success(message)
                    else:
                        st.sidebar.error(message)
        else:
            st.sidebar.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
    else:
        st.sidebar.error("USER NAME.xlsx file not found in the repository.")

    # Download option
    if st.sidebar.button("📂 Download Data"):
        timestamp = datetime.now().strftime("%d%m%y%H%M%S")
        file_name = f"UnauthorizedAccess_{timestamp}.xlsx"
        excel_data = convert_df_to_excel_with_sheets(mismatches_df, rms_df, current_alarms_df, site_access_df)
        
        st.sidebar.download_button(
            label="Click to download",
            data=excel_data,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Telegram Notification Option
    if st.sidebar.button("💬 Send Notification"):
        if os.path.exists(user_file_path):
            user_df = pd.read_excel(user_file_path)
            user_df = clean_column_names(user_df)
            user_df['Name'] = user_df['Name'].astype(str)  # Ensure names are strings

            if "Zone" in user_df.columns and "Name" in user_df.columns:
                zone_to_name = user_df.set_index("Zone")["Name"].to_dict()
                zones = filtered_mismatches_df['Zone'].unique()
                bot_token = "7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo"
                chat_id = "-4625672098"

                for zone in zones:
                    zone_df = filtered_mismatches_df[filtered_mismatches_df['Zone'] == zone]
                    zone_df['End Time'] = zone_df['End Time'].replace("Not Closed", None)
                    sorted_zone_df = zone_df.sort_values(by='End Time', na_position='first')
                    sorted_zone_df['End Time'] = sorted_zone_df['End Time'].fillna("Not Closed")

                    message = f"❗Door Open Notification❗\n\n🚩 {zone}\n\n"
                    site_aliases = sorted_zone_df['Site Alias'].unique()

                    for site_alias in site_aliases:
                        site_df = sorted_zone_df[sorted_zone_df['Site Alias'] == site_alias]
                        message += f"✔ {site_alias}\n"
                        for _, row in site_df.iterrows():
                            end_time_display = row['End Time']
                            message += f"  • Start Time: {row['Start Time']} | End Time: {end_time_display}\n"
                        message += "\n"

                    # Handle the name safely
                    zone_name = zone_to_name.get(zone, "")
                    if pd.isna(zone_name):
                        zone_name = ""
                    escaped_name = str(zone_name).replace("_", "\\_")
                    message += f"**@{escaped_name}**, no Site Access Request found for these Door Open alarms. Please take care and share us update.\n"

                    payload = {
                        "chat_id": chat_id,
                        "text": message,
                        "parse_mode": "Markdown"
                    }
                    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
                    response = requests.post(url, json=payload)

                    if response.status_code == 200:
                        st.success(f"Notification for zone '{zone}' sent successfully!")
                    else:
                        st.error(f"Failed to send notification for zone '{zone}'.")
            else:
                st.error("The USER NAME.xlsx file must have 'Zone' and 'Name' columns.")
        else:
            st.error("USER NAME.xlsx file not found in the repository.")
