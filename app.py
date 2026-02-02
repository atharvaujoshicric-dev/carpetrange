import streamlit as st
import pandas as pd
import numpy as np

def process_data(df):
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Column Names
    ft_col = 'Carpet Area(SQ.FT)'
    min_apr = 'Min. APR'
    max_apr = 'Max APR'
    avg_apr = 'Average of APR'
    med_apr = 'Median of APR'

    # 1. Convert to numeric and remove decimals (Rounding to 0)
    cols_to_clean = [ft_col, min_apr, max_apr, avg_apr, med_apr]
    for col in cols_to_clean:
        if col in df.columns:
            # pd.to_numeric handles strings, .round(0) removes decimals
            df[col] = pd.to_numeric(df[col], errors='coerce').round(0)

    # 2. Create ranges in multiples of 50
    if ft_col in df.columns:
        # Calculate the lower bound for each area (e.g., 430 becomes 400, 460 becomes 450)
        # We use floor division to get the bucket
        df['Range_Start'] = (df[ft_col] // 50) * 50
        df['Range_End'] = df['Range_Start'] + 50
        
        # Create a string label: "400-450"
        df['Carpet Area Range'] = (
            df['Range_Start'].astype(int).astype(str) + 
            "-" + 
            df['Range_End'].astype(int).astype(str)
        )

        # 3. Group by the new Range
        grouped = df.groupby('Carpet Area Range').agg({
            min_apr: 'min',
            max_apr: 'max',
            avg_apr: 'mean'
        }).reset_index()

        # Round the final grouped average and remove decimals
        grouped[avg_apr] = grouped[avg_apr].round(0)
    else:
        st.error(f"Column '{ft_col}' not found.")
        return df, None

    # 4. Remove Median column and helper range columns
    if med_apr in df.columns:
        df = df.drop(columns=[med_apr])
    
    return df, grouped

# Streamlit UI
st.set_page_config(page_title="Real Estate Tool", layout="wide")
st.title("Property Summary: 50 Sq.Ft Range Logic")

uploaded_file = st.file_uploader("Upload Summary Excel/CSV", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')

        processed_df, summary_df = process_data(df)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Cleaned Data (No Decimals)")
            st.dataframe(processed_df.drop(columns=['Range_Start', 'Range_End'], errors='ignore'))

        with col2:
            st.subheader("Grouped by 50 Sq.Ft Multiples")
            if summary_df is not None:
                st.table(summary_df)
                
    except Exception as e:
        st.error(f"Error: {e}. Ensure columns are named correctly.")
