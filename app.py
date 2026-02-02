import streamlit as st
import pandas as pd

def process_data(df):
    # Clean column names to prevent "not found" errors
    df.columns = df.columns.str.strip()
    
    # Define your target columns
    ft_col = 'Carpet Area(SQ.FT)'
    min_apr = 'Min. APR'
    max_apr = 'Max APR'
    avg_apr = 'Average of APR'
    med_apr = 'Median of APR'

    # 1. Round off decimals in specific columns
    cols_to_round = [ft_col, min_apr, max_apr, avg_apr, med_apr]
    for col in cols_to_round:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

    # 2. Group rows by the SQ.FT column
    if ft_col in df.columns:
        # We group by the area and calculate the stats for the prices
        grouped = df.groupby(ft_col).agg({
            min_apr: 'min',
            max_apr: 'max',
            avg_apr: 'mean'
        }).reset_index()
        
        # Round the resulting average calculation
        grouped[avg_apr] = grouped[avg_apr].round(2)
    else:
        st.error(f"Column '{ft_col}' not found. Please check the Excel headers.")
        return df, None

    # 3. Remove Median APR column from the main view
    if med_apr in df.columns:
        df = df.drop(columns=[med_apr])
        
    return df, grouped

# Streamlit UI
st.set_page_config(page_title="Real Estate Tool", layout="wide")
st.title("Property Summary Processor")

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
            st.subheader("Cleaned Data (Median Removed)")
            st.dataframe(processed_df, use_container_width=True)

        with col2:
            st.subheader("Grouped by Carpet Area (SQ.FT)")
            if summary_df is not None:
                st.dataframe(summary_df, use_container_width=True)
                
    except Exception as e:
        st.error(f"Error: {e}. Ensure the sheet is named 'Summary'.")
