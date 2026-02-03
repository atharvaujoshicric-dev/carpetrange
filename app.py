import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Real Estate Report Generator", layout="wide")

st.title("🏙️ Real Estate Report Generator")
st.write("Upload your Excel/CSV file with a 'summary' sheet to generate the 'report' sheet.")

# 1. File Uploader
uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # 2. Read the data
        # If CSV, we assume it's the summary. If Excel, we look for the 'summary' sheet.
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        st.success("File uploaded successfully! Processing data...")

        # 3. Data Transformation Logic
        # Group by Property and Configuration
        report_df = df.groupby(['Property', 'Configuration']).agg({
            'Carpet Area(SQ.FT)': ['min', 'max'],
            'Average of APR': 'mean',
            'Count of Property': 'sum',
            'Total Count': 'sum'
        }).reset_index()

        # Rename columns to match your requirements
        report_df.columns = [
            'Property', 'Configurations', 'Min. Carpet Area', 
            'Max. Carpet Area', 'Avg APR', 'Count of Property', 'Total Count'
        ]

        # 4. Display the result in the app
        st.subheader("Preview: New Report Sheet")
        st.dataframe(report_df)

        # 5. Prepare Excel file for download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Keep the original summary if it's an Excel file
            df.to_excel(writer, index=False, sheet_name='summary')
            # Add the new report sheet
            report_df.to_excel(writer, index=False, sheet_name='report')
        
        processed_data = output.getvalue()

        # 6. Download Button
        st.download_button(
            label="📥 Download Updated Excel File",
            data=processed_data,
            file_name="Property_Report_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Make sure your file has a sheet named 'summary' and the correct column names.")

else:
    st.info("Waiting for file upload...")
