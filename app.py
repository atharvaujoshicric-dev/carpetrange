import streamlit as st
import pandas as pd
import io

def process_data(df):
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Mapping exact column names from your request
    prop_col = 'Property'
    config_col = 'Configuration'
    carpet_col = 'Carpet Area(SQ.FT)'
    apr_col = 'Average of APR'
    
    # Check if required columns exist
    required = [prop_col, config_col, carpet_col, apr_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Missing columns in 'Summary' sheet: {missing}")
        return None

    # 1. Ensure numeric data and round off (remove decimals)
    df[carpet_col] = pd.to_numeric(df[carpet_col], errors='coerce')
    df[apr_col] = pd.to_numeric(df[apr_col], errors='coerce')

    # 2. Grouping Logic: Combine similar configs for each property
    # We calculate: Min Carpet, Max Carpet, Avg APR, and Total Count
    summary_report = df.groupby([prop_col, config_col]).agg(
        Min_Carpet=(carpet_col, 'min'),
        Max_Carpet=(carpet_col, 'max'),
        Avg_APR=(apr_col, 'mean'),
        Total_Count=(config_col, 'count')
    ).reset_index()

    # 3. Final Rounding (Removing decimals)
    summary_report['Min_Carpet'] = summary_report['Min_Carpet'].round(0)
    summary_report['Max_Carpet'] = summary_report['Max_Carpet'].round(0)
    summary_report['Avg_APR'] = summary_report['Avg_APR'].round(0)
    summary_report['Total_Count'] = summary_report['Total_Count'].astype(int)

    return summary_report

# Streamlit UI Setup
st.set_page_config(page_title="Project Config Reporter", layout="wide")
st.title("Real Estate Project Configuration Summary")

uploaded_file = st.file_uploader("Upload your Excel/CSV file", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        report_df = process_data(df)

        if report_df is not None:
            st.subheader("Property-wise Configuration Report")
            st.dataframe(report_df, use_container_width=True)

            # Create Excel Download in memory
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                report_df.to_excel(writer, index=False, sheet_name='Config Summary')
            
            st.download_button(
                label="📥 Download New Sheet as Excel",
                data=buffer.getvalue(),
                file_name="Project_Configuration_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
                
    except Exception as e:
        st.error(f"Error: {e}. Please ensure the sheet 'Summary' exists and column names match.")
