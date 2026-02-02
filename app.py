import streamlit as st
import pandas as pd
import io

def process_data(df):
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Define column names
    prop_col = 'Property'
    config_col = 'Configuration'
    carpet_col = 'Carpet Area(SQ.FT)'
    apr_col = 'Average of APR'
    
    # Validate columns
    required = [prop_col, config_col, carpet_col, apr_col]
    if not all(col in df.columns for col in required):
        st.error(f"Required columns missing. Found: {list(df.columns)}")
        return None

    # Ensure numeric for calculations
    df[carpet_col] = pd.to_numeric(df[carpet_col], errors='coerce')
    df[apr_col] = pd.to_numeric(df[apr_col], errors='coerce')

    # Group by Property then Configuration
    # We sum the APRs and count the units to do the custom division
    summary = df.groupby([prop_col, config_col]).agg(
        Min_Carpet=(carpet_col, 'min'),
        Max_Carpet=(carpet_col, 'max'),
        Sum_APR=(apr_col, 'sum'),
        Total_Count=(config_col, 'count')
    ).reset_index()

    # Apply your specific formula: Sum of APRs / Total Count
    summary['Calculated Avg APR'] = (summary['Sum_APR'] / summary['Total_Count']).round(0)
    
    # Round Carpet areas
    summary['Min_Carpet'] = summary['Min_Carpet'].round(0)
    summary['Max_Carpet'] = summary['Max_Carpet'].round(0)

    # Reorganize columns for the final report
    final_report = summary[[
        prop_col, 
        config_col, 
        'Min_Carpet', 
        'Max_Carpet', 
        'Calculated Avg APR', 
        'Total_Count'
    ]]
    
    return final_report

# Streamlit UI
st.set_page_config(page_title="Real Estate Analytics", layout="wide")
st.title("Project Summary & Configuration Analysis")

uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        report = process_data(df)

        if report is not None:
            # Sorting so Property names stay together
            report = report.sort_values(by=['Property', 'Configuration'])
            
            st.subheader("Final Configuration Summary")
            st.dataframe(report, use_container_width=True, hide_index=True)

            # Excel Export logic
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                report.to_excel(writer, index=False, sheet_name='Config Summary')
            
            st.download_button(
                label="📥 Download Summary Report",
                data=buffer.getvalue(),
                file_name="Property_Config_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
                
    except Exception as e:
        st.error(f"Error: {e}. Ensure the 'Summary' sheet exists.")
