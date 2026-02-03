import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Property Report Tool", layout="wide")

st.title("🏙️ Real Estate Summary to Report")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # 1. Load the summary sheet
        df = pd.read_excel(uploaded_file, sheet_name='summary')

        # 2. Handle Merged Cells: Fill Property and Total Count downwards
        df['Property'] = df['Property'].ffill()
        df['Total Count'] = df['Total Count'].ffill()

        # 3. Grouping Logic to match your target image
        # We group by Property, Last Completion Date, and Configuration
        report_df = df.groupby(['Property', 'Last Completion Date', 'Configuration']).agg({
            'Carpet Area(SQ.FT)': ['min', 'max'],
            'Average of APR': 'mean',
            'Count of Property': 'sum',
            'Total Count': 'first'
        }).reset_index()

        # 4. Flatten columns and rename to match your target screenshot
        report_df.columns = [
            'Property', 'Last Completion Date', 'Configuration', 
            'Min. Carpet Area(SQ.FT)', 'Max. Carpet Area(SQ.FT)', 
            'Average of APR', 'Count of Property', 'Total Count'
        ]

        # Formatting the date to match 'Aug-28' style if needed
        report_df['Last Completion Date'] = pd.to_datetime(report_df['Last Completion Date']).dt.strftime('%b-%y')

        st.subheader("Preview of Generated Report")
        st.dataframe(report_df)

        # 5. Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='summary')
            report_df.to_excel(writer, index=False, sheet_name='report')
        
        st.download_button(
            label="📥 Download Updated Excel",
            data=output.getvalue(),
            file_name="Property_Report_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
