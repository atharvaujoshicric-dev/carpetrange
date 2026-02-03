import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Property Report Tool", layout="wide")

st.title("🏙️ Real Estate Summary to Report")
st.write("Calculates weighted averages, combines carpet areas, and applies custom styling.")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # 1. Handle flexible sheet naming (summary or Summary)
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = next((s for s in xls.sheet_names if s.lower() == 'summary'), None)
        
        if not target_sheet:
            st.error("Could not find a sheet named 'summary' or 'Summary'.")
        else:
            df = pd.read_excel(xls, sheet_name=target_sheet)

            # 2. Fix Merged Cells & Date Formats
            df['Property'] = df['Property'].ffill()
            df['Total Count'] = df['Total Count'].ffill()
            df['Last Completion Date'] = pd.to_datetime(df['Last Completion Date'], dayfirst=True)

            # Weighted sum for accurate APR
            df['weighted_apr_sum'] = df['Average of APR'] * df['Count of Property']

            # 3. Aggregation (Grouping by Property + Date + Config)
            report_df = df.groupby(['Property', 'Last Completion Date', 'Configuration']).agg({
                'Carpet Area(SQ.FT)': ['min', 'max'],
                'Min. APR': 'min',
                'Max APR': 'max',
                'weighted_apr_sum': 'sum',
                'Count of Property': 'sum',
                'Total Count': 'first'
            }).reset_index()

            # Flatten Multi-index columns
            report_df.columns = [
                'Property', 'Last Completion Date', 'Configuration', 
                'min_carpet', 'max_carpet', 'Min APR', 'Max APR', 
                'temp_sum_apr', 'Count of Property', 'Total Count'
            ]

            # 4. Data Formatting
            # Weighted Average calculation
            report_df['Average of APR'] = report_df['temp_sum_apr'] / report_df['Count of Property']
            
            # Combine Carpet Area into "min-max"
            report_df['Carpet Area(SQ.FT)'] = (
                report_df['min_carpet'].round(0).astype(int).astype(str) + "-" + 
                report_df['max_carpet'].round(0).astype(int).astype(str)
            )

            # Round other numeric columns to 0 decimals
            round_cols = ['Min APR', 'Max APR', 'Average of APR', 'Count of Property', 'Total Count']
            report_df[round_cols] = report_df[round_cols].round(0).astype(int)

            # Format Date
            report_df['Last Completion Date'] = report_df['Last Completion Date'].dt.strftime('%b-%y')

            # Final Column Order
            final_df = report_df[['Property', 'Last Completion Date', 'Configuration', 
                                  'Carpet Area(SQ.FT)', 'Min APR', 'Max APR', 
                                  'Average of APR', 'Count of Property', 'Total Count']]

            st.subheader("Preview of Styled Report")
            st.dataframe(final_df)

            # 5. Advanced Excel Export with Styling
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write original summary
                df.drop(columns=['weighted_apr_sum']).to_excel(writer, index=False, sheet_name='summary')
                # Write report
                final_df.to_excel(writer, index=False, sheet_name='report')
                
                # Get the worksheet for styling
                workbook = writer.book
                worksheet = workbook['report']

                # --- Excel Styling (Merging and Colors) ---
                # Unique Colors for each property (Hex codes)
                colors = ['FFDDC1', 'C1E1C1', 'C1D4E1', 'E1C1E1', 'FDFD96', 'FFB7B2', 'B2CEFE']
                property_colors = {}
                
                # Group properties to find merge ranges
                current_prop = None
                start_row = 2
                color_idx = 0
                
                last_row = len(final_df) + 1
                
                for row_num in range(2, last_row + 2):
                    row_prop = worksheet.cell(row=row_num, column=1).value
                    
                    if row_prop != current_prop or row_num == last_row + 1:
                        if current_prop is not None:
                            end_row = row_num - 1
                            
                            # Apply unique color to this property block
                            fill_color = PatternFill(start_color=colors[color_idx % len(colors)], 
                                                     end_color=colors[color_idx % len(colors)], fill_type="solid")
                            
                            for r in range(start_row, end_row + 1):
                                for c in range(1, 10): # Cols A to I
                                    worksheet.cell(row=r, column=c).fill = fill_color
                            
                            # Merge Property and Total Count
                            if end_row > start_row:
                                worksheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                                worksheet.merge_cells(start_row=start_row, start_column=9, end_row=end_row, end_column=9)
                            
                            # Center the merged cells
                            worksheet.cell(row=start_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
                            worksheet.cell(row=start_row, column=9).alignment = Alignment(horizontal='center', vertical='center')
                            
                            color_idx += 1
                        
                        start_row = row_num
                        current_prop = row_prop

                # Auto-adjust column widths
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    worksheet.column_dimensions[column].width = max_length + 2

            st.download_button(
                label="📥 Download Styled Excel Report",
                data=output.getvalue(),
                file_name="Property_Report_Styled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
