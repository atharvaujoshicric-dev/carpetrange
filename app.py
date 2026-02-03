import pandas as pd

def generate_report(file_path):
    # Load the summary sheet
    df = pd.read_excel(file_path, sheet_name='summary')

    # Group by Property AND Configuration to get Min/Max Carpet Area
    report_df = df.groupby(['Property', 'Configuration']).agg({
        'Carpet Area(SQ.FT)': ['min', 'max'],
        'Average of APR': 'first', # Or 'mean' depending on your data
        'Count of Property': 'first',
        'Total Count': 'first'
    }).reset_index()

    # Flatten the columns (since groupby creates a MultiIndex)
    report_df.columns = [
        'Property', 'Configurations', 'Min. Carpet Area', 
        'Max. Carpet Area', 'Avg APR', 'Count of Property', 'Total Count'
    ]
    
    return report_df
