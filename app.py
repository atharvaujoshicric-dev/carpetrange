import streamlit as st
import pandas as pd
import numpy as np

def process_data(df):
    # 1. Round off decimals in specific columns
    cols_to_round = ['Carpet Area(SQ.FT)', 'Min. APR', 'Max APR', 'Average of APR', 'Median of APR']
    
    # Ensure columns exist before rounding to avoid errors
    existing_cols = [col for col in cols_to_round if col in df.columns]
    df[existing_cols] = df[existing_cols].round(2)

    # 2. Define the Carpet Area (SQ.MT) ranges
    # You can adjust the bins/labels based on your specific project needs
    if 'Carpet Area (SQ.MT)' in df.columns:
        bins = [0, 40, 60, 80, 100, 150, 500]
        labels = ['0-40', '40-60', '60-80', '80-100', '100-150', '150+']
        df['Area Range (SQ.MT)'] = pd.cut(df['Carpet Area (SQ.MT)'], bins=bins, labels=labels)

        # 3. Grouping logic
        grouped = df.groupby('Area Range (SQ.MT)').agg({
            'Min. APR': 'min',
            'Max APR': 'max',
            'Average of APR': 'mean'
        }).reset_index()
        
        # Round the new average calculation
        grouped['Average of APR'] = grouped['Average of APR'].round(2)
    else:
        st.error("Column 'Carpet Area (SQ.MT)' not found.")
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
