import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Property Report Tool", layout="wide")

st.title("🏙️ Real Estate Summary to Report")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = next((s for s in xls.sheet_names if s.lower() == 'summary'), None)
        
        if not target_sheet:
            st.error("Could not find a sheet named 'summary' or 'Summary'.")
        else:
            df = pd.read_excel(xls, sheet_name=target_sheet)

            # 1. Fix Merged Cells & Date Formats
            df['Property'] = df['Property'].ffill()
            df['Total Count'] = df['Total Count'].ffill()
            df['Last Completion Date'] = pd.to_datetime(df['Last Completion Date'], dayfirst=True)
            df['weighted_apr_sum'] = df['Average of APR'] * df['Count of Property']

            # 2. Aggregation
            report_df = df.groupby(['Property', 'Last Completion Date', 'Configuration']).agg({
                'Carpet Area(SQ.FT)': ['min', 'max'],
                'Min. APR': 'min',
                'Max APR': 'max',
                'weighted_apr_sum': 'sum',
                'Count of Property': 'sum',
                'Total Count': 'first'
            }).reset_index()

            report_df.columns = [
                'Property', 'Last Completion Date', 'Configuration', 
                'min_carpet', 'max_carpet', 'Min APR', 'Max APR', 
                'temp_sum_apr', 'Count of Property', 'Total Count'
            ]

            # 3. Logic for Carpet Area (Single value if Min == Max)
            def format_carpet(row):
                mi, ma = round(row['min_carpet']), round(row['max_carpet'])
                return str(int(mi)) if mi == ma else f"{int(mi)}-{int(ma)}"

            report_df['Carpet Area(SQ.FT)'] = report_df.apply(format_carpet, axis=1)

            # 4. Weighted Average and Rounding
            report_df['Average of APR'] = (report_df['temp_sum_apr'] / report_df['Count of Property']).round(0).astype(int)
            round_cols = ['Min APR', 'Max APR', 'Count of Property', 'Total Count']
            report_df[round_cols] = report_df[round_cols].round(0).astype(int)
            report_df['Last Completion Date'] = report_df['Last Completion Date'].dt.strftime('%b-%y')

            final_df = report_df[['Property', 'Last Completion Date', 'Configuration', 
                                  'Carpet Area(SQ.FT)', 'Min APR', 'Max APR', 
                                  'Average of APR', 'Count of Property', 'Total Count']]

            st.subheader("Preview")
            st.dataframe(final_df)

            # 5. Advanced Excel Export with Borders and Styling
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='report')
                workbook = writer.book
                worksheet = workbook['report']

                # Define Border Style
                thin_border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin')
                )

                # Unique Colors for Properties
                colors = ['FFDDC1', 'C1E1C1', 'C1D4E1', 'E1C1E1', 'FDFD96', 'FFB7B2', 'B2CEFE']
                current_prop = None
                start_row = 2
                color_idx = 0
                last_row = len(final_df) + 1
                
                for row_num in range(2, last_row + 2):
                    # Apply borders to every cell in the range
                    if row_num <= last_row:
                        for col_num in range(1, 10):
                            worksheet.cell(row=row_num, column=col_num).border = thin_border

                    row_prop = worksheet.cell(row=row_num, column=1).value
                    
                    if row_prop != current_prop or row_num == last_row + 1:
                        if current_prop is not None:
                            end_row = row_num - 1
                            fill = PatternFill(start_color=colors[color_idx % len(colors)], fill_type="solid")
                            
                            for r in range(start_row, end_row + 1):
                                for c in range(1, 10):
                                    worksheet.cell(row=r, column=c).fill = fill
                            
                            if end_row > start_row:
                                worksheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                                worksheet.merge_cells(start_row=start_row, start_column=9, end_row=end_row, end_column=9)
                            
                            worksheet.cell(row=start_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
                            worksheet.cell(row=start_row, column=9).alignment = Alignment(horizontal='center', vertical='center')
                            color_idx += 1
                        
                        start_row = row_num
                        current_prop = row_prop

                # Style Header Row
                for cell in worksheet[1]:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center')

            st.download_button(
                label="📥 Download Bordered & Styled Report",
                data=output.getvalue(),
                file_name="Property_Report_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
