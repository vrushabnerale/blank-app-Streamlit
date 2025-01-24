import streamlit as st
import pandas as pd
import re

# Function to extract the 5/6-digit code from Prozessname
def extract_code(prozess_name):
    pattern = r'\b(HDD|KV|OBW)[-\s]\d{2,3}-\d{2}\b'
    match = re.search(pattern, str(prozess_name))
    return match.group(0) if match else None

# Streamlit App
st.title("Excel Data Filtering and Processing App")

# File Upload
uploaded_files = st.file_uploader("Upload Excel Files", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    # Load files and process them
    st.write("Uploaded Files:")
    for file in uploaded_files:
        st.write(f"- {file.name}")

    # User-defined cutoff for KW Start
    kw_cutoff = st.number_input("KW Cutoff (e.g., 27)", min_value=1, max_value=52, value=27, step=1)

    # Process the files
    filtered_dataframes = []

    for file in uploaded_files:
        df = pd.read_excel(file)

        # Ensure Startdatum is in datetime format
        df['Startdatum'] = pd.to_datetime(df['Startdatum'], errors='coerce')

        # Filter rows where Startdatum is in the year 2025
        filtered_df = df[df['Startdatum'].dt.year == 2025].copy()

        # Filter for KW Start
        KW_cutoff_df = filtered_df[filtered_df['KW Start'] <= kw_cutoff].copy()

        # Filter Gewerke for specific keywords
        keywords = ['HDD', 'OBW', 'offene Bauweise', 'Kurzvortrieb', 'Mikrotunnel', 'MT']
        pattern = '|'.join(keywords)
        gewerk_filtered_df = KW_cutoff_df[KW_cutoff_df['Gewerk'].str.contains(pattern, case=False, na=False)].copy()

        # Filter Prozessname for specific includes and excludes
        includes = ['Fertigstellung', 'OBW', 'HDD']
        excludes = [
            'Vorarbeit', 'Zuwegung', 'PE-Rohre Schweißen', 'PE-Schweißen',
            'Anbindung', 'Oberboden auftragen', 'Oberbodenauftrag',
            'obw baustraße', 'Deichkreuzung', 'Teil'
        ]

        include_pattern = '|'.join(includes)
        exclude_pattern = '|'.join(excludes)

        prozessname_filtered_df = gewerk_filtered_df[
            gewerk_filtered_df['Prozessname'].str.contains(include_pattern, case=False, na=False) &
            ~gewerk_filtered_df['Prozessname'].str.contains(exclude_pattern, case=False, na=False)
        ].copy()

        # Add columns
        prozessname_filtered_df['NDS/NRW'] = file.name.split('.')[0]
        prozessname_filtered_df['Bauweise\nBereichs-\nerkennung'] = prozessname_filtered_df['Prozessname'].apply(extract_code)

        # Keep specific columns
        columns_to_keep = ['Id', 'Prozessname', 'Startdatum', 'Enddatum', 'Status', 'Dauer', 'Gewerk', 'KW Start', 'KW Ende', 'NDS/NRW', 'Bauweise\nBereichs-\nerkennung']
        prozessname_filtered_df = prozessname_filtered_df[columns_to_keep]

        filtered_dataframes.append(prozessname_filtered_df)

    # Combine all filtered DataFrames
    combined_df = pd.concat(filtered_dataframes, ignore_index=True)

    # Display preview of the combined data
    st.subheader("Preview of Processed Data")
    st.dataframe(combined_df.head(20))

     # Analysis 1: Count of Entries by NDS/NRW
    st.subheader("Count of Entries by NDS/NRW")
    counts = combined_df['NDS/NRW'].value_counts()
    st.table(counts)
    st.bar_chart(counts)

    # Option to download the processed data
    output_file = "combined_filtered_data.xlsx"
    combined_df.to_excel(output_file, index=False)
    st.download_button(
        label="Download Processed Data",
        data=open(output_file, 'rb').read(),
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
