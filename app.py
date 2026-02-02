import streamlit as st
import pandas as pd
import io

def process_data(df):
    # Clean column names to handle any hidden spaces
    df.columns = df.columns.str.strip()
    
    # Required Column Mapping
    prop_col = 'Property'
    config_col = 'Configuration'
    carpet_col = 'Carpet Area(SQ.FT)'
    apr_col = 'Average of APR'
    
    # Validation
    required = [prop_col, config_col, carpet_col, apr_col]
    if not all(col in df.columns for col in required):
        st.error(f"Please check column names. Need: {required}")
        return None

    # Ensure numeric types for calculation
    df[carpet_col] = pd.to_numeric(df[carpet_col], errors='coerce')
    df[apr_col] = pd.to_numeric(df[apr_col], errors='coerce')

    # Grouping by Property, then Configuration
    # We calculate Min Carpet, Sum of APR, and the Count (Total units)
    summary = df.groupby([prop_col, config_col]).agg(
        Min_Carpet=(carpet_col, 'min'),
        Total_APR_Sum=(apr_col, 'sum'),
        Unit_Count=(config_col, 'count')
    ).reset_index()

    # Calculate Manual Avg APR: (Sum of APRs / Total Count)
    summary['Calculated Avg APR'] = summary['Total_APR_Sum'] / summary['Unit_Count']

    # Rounding and Formatting (Removing Decimals)
    summary['Min_Carpet'] = summary['Min_Carpet'].round(0).astype(int)
    summary['Calculated Avg APR'] = summary['Calculated Avg APR'].round(0).astype(int)
    
    # Final cleanup: removing the helper sum column before showing the user
    final_report = summary[[prop_col, config_col, 'Min_Carpet', 'Calculated Avg APR', 'Unit_Count']]
    
    return final_report

# Streamlit Interface
st.set_page_config(page_title="Real Estate Summary Tool", layout="wide")
st.title("Property & Configuration Summarizer")

uploaded_file = st.file_uploader("Upload Summary Excel", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        result_df = process_data(df)

        if result_df is not None:
            st.subheader("Summarized Report")
            # Displaying the data
            st.dataframe(result_df, use_container_width=True)

            # Export Logic
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Property Report')
            
            st.download_button(
                label="📥 Download Summarized Sheet",
                data=buffer.getvalue(),
                file_name="Property_Config_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
                
    except Exception as e:
        st.error(f"Error processing file: {e}")
