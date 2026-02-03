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

        # 2. Fix Merged Cells & Date Formats
        df['Property'] = df['Property'].ffill()
        df['Total Count'] = df['Total Count'].ffill()
        
        # Convert date safely - handles dayfirst for DD-MM-YYYY
        df['Last Completion Date'] = pd.to_datetime(df['Last Completion Date'], dayfirst=True)

        # 3. Weighted Average Calculation for APR (to get that 5308.5 logic correct)
        df['weighted_apr_sum'] = df['Average of APR'] * df['Count of Property']

        # 4. Grouping Logic
        report_df = df.groupby(['Property', 'Last Completion Date', 'Configuration']).agg({
            'Carpet Area(SQ.FT)': ['min', 'max'],
            'weighted_apr_sum': 'sum',
            'Count of Property': 'sum',
            'Total Count': 'first'
        }).reset_index()

        # 5. Flatten columns
        report_df.columns = [
            'Property', 'Last Completion Date', 'Configuration', 
            'Min. Carpet Area(SQ.FT)', 'Max. Carpet Area(SQ.FT)', 
            'temp_sum_apr', 'Count of Property', 'Total Count'
        ]
        
        # Calculate Weighted Average
        report_df['Average of APR'] = report_df['temp_sum_apr'] / report_df['Count of Property']

        # --- ROUNDING LOGIC ---
        # Rounding all numeric columns to 0 decimal places
        cols_to_round = [
            'Min. Carpet Area(SQ.FT)', 'Max. Carpet Area(SQ.FT)', 
            'Average of APR', 'Count of Property', 'Total Count'
        ]
        report_df[cols_to_round] = report_df[cols_to_round].round(0).astype(int)

        # Clean up and Format Date to "Aug-28"
        report_df['Last Completion Date'] = report_df['Last Completion Date'].dt.strftime('%b-%y')
        final_df = report_df[['Property', 'Last Completion Date', 'Configuration', 
                              'Min. Carpet Area(SQ.FT)', 'Max. Carpet Area(SQ.FT)', 
                              'Average of APR', 'Count of Property', 'Total Count']]

        st.subheader("Preview of Generated Report (Rounded)")
        st.dataframe(final_df)

        # 6. Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # We keep the original summary as is, or you can round it too if you like
            df.drop(columns=['weighted_apr_sum']).to_excel(writer, index=False, sheet_name='summary')
            final_df.to_excel(writer, index=False, sheet_name='report')
        
        st.download_button(
            label="📥 Download Rounded Report",
            data=output.getvalue(),
            file_name="Property_Report_Rounded.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
