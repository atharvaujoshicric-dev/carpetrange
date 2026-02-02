import streamlit as st
import pandas as pd
import numpy as np

def process_data(df):
    # Clean column names: remove leading/trailing spaces and handle internal spacing issues
    df.columns = df.columns.str.strip()
    
    # 1. Round off decimals in specific columns
    # Using a flexible list to match your exact naming
    cols_to_round = ['Carpet Area(SQ.FT)', 'Min. APR', 'Max APR', 'Average of APR', 'Median of APR']
    
    # Check for columns and round them if they exist
    for col in cols_to_round:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

    # 2. Define the Carpet Area (SQ.MT) ranges
    # Standardizing the name for Carpet Area (SQ.MT)
    sq_mt_col = 'Carpet Area (SQ.MT)'
    
    if sq_mt_col in df.columns:
        # Converting to numeric just in case there are strings/nulls
        df[sq_mt_col] = pd.to_numeric(df[sq_mt_col], errors='coerce')
        
        bins = [0, 40, 60, 80, 100, 150, 1000]
        labels = ['0-40', '40-60', '60-80', '80-100', '100-150', '150+']
        df['Area Range (SQ.MT)'] = pd.cut(df[sq_mt_col], bins=bins, labels=labels)

        # 3. Grouping logic
        grouped = df.groupby('Area Range (SQ.MT)', observed=False).agg({
            'Min. APR': 'min',
            'Max APR': 'max',
            'Average of APR': 'mean'
        }).reset_index()
        
        grouped['Average of APR'] = grouped['Average of APR'].round(2)
    else:
        st.error(f"Column '{sq_mt_col}' not found. Available columns are: {list(df.columns)}")
        return df, None

    # 4. Remove Median APR column
    if 'Median of APR' in df.columns:
        df = df.drop(columns=['Median of APR'])
        
    return df, grouped

# Streamlit UI
st.title("Real Estate Summary Processor")

uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # specifically targeting the 'Summary' sheet
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        processed_df, summary_df = process_data(df)

        st.subheader("Processed Data")
        st.dataframe(processed_df)

        if summary_df is not None:
            st.subheader("Grouped Area Analysis")
            st.table(summary_df)
            
    except Exception as e:
        st.error(f"Error: {e}. Please ensure the sheet 'Summary' exists and columns are named correctly.")
