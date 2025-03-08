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

# Streamlit app
st.title('🛡️IntrusionShield🛡️')

# File uploaders
site_access_file = st.file_uploader("Upload the Site Access Data", type=["xlsx"])
rms_file = st.file_uploader("Upload the All Door Open Alarms Data till now", type=["xlsx"])
current_alarms_file = st.file_uploader("Upload the Current Door Open Alarms Data", type=["xlsx"])

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
            st.write(mismatches_df)
        else:
            st.write("No mismatched sites found.")
    else:
        st.error("Merged DataFrame is empty. Check the input files and column names.")
