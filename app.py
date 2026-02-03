import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Property Report Tool", layout="wide")

st.title("🏙️Spydarr's Summary to Report")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = next((s for s in xls.sheet_names if s.lower() == 'summary'), None)
        
        if not target_sheet:
            st.error("Could not find a sheet named 'summary' or 'Summary'.")
        else:
            df = pd.read_excel(xls, sheet_name=target_sheet)

            # 1. Data Cleaning & Calculations
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

            # 3. Logic for Carpet Area & Rounding
            report_df['Carpet Area(SQ.FT)'] = report_df.apply(
                lambda x: str(int(round(x['min_carpet']))) if round(x['min_carpet']) == round(x['max_carpet']) 
                else f"{int(round(x['min_carpet']))}-{int(round(x['max_carpet']))}", axis=1
            )

            report_df['Average of APR'] = (report_df['temp_sum_apr'] / report_df['Count of Property']).round(0).astype(int)
            for col in ['Min APR', 'Max APR', 'Count of Property', 'Total Count']:
                report_df[col] = report_df[col].round(0).astype(int)
            
            report_df['Last Completion Date'] = report_df['Last Completion Date'].dt.strftime('%b-%y')

            final_df = report_df[['Property', 'Last Completion Date', 'Configuration', 
                                  'Carpet Area(SQ.FT)', 'Min APR', 'Max APR', 
                                  'Average of APR', 'Count of Property', 'Total Count']]

            st.subheader("Styled Preview")
            st.dataframe(final_df)

            # 4. Excel Styling with Center/Middle Alignment
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='report')
                workbook = writer.book
                ws = workbook['report']

                # Alignment & Border Styles
                center_alignment = Alignment(horizontal='center', vertical='center')
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                
                colors = ['FFDDC1', 'C1E1C1', 'C1D4E1', 'E1C1E1', 'FDFD96', 'FFB7B2', 'B2CEFE']
                current_prop, start_row, color_idx = None, 2, 0
                last_row = len(final_df) + 1

                # Loop through all data cells for alignment and borders
                for r in range(1, last_row + 1):
                    for c in range(1, 10):
                        cell = ws.cell(row=r, column=c)
                        cell.alignment = center_alignment
                        cell.border = thin_border

                # Handle Merging and Property Coloring
                for row_num in range(2, last_row + 2):
                    row_prop = ws.cell(row=row_num, column=1).value
                    if row_prop != current_prop or row_num == last_row + 1:
                        if current_prop is not None:
                            end_row = row_num - 1
                            fill = PatternFill(start_color=colors[color_idx % len(colors)], fill_type="solid")
                            
                            for r_fill in range(start_row, end_row + 1):
                                for c_fill in range(1, 10):
                                    ws.cell(row=r_fill, column=c_fill).fill = fill
                            
                            if end_row > start_row:
                                ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                                ws.merge_cells(start_row=start_row, start_column=9, end_row=end_row, end_column=9)
                            
                            color_idx += 1
                        start_row, current_prop = row_num, row_prop

                # Auto-adjust column width
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 18

            st.download_button(
                label="📥 Download Centered Report",
                data=output.getvalue(),
                file_name="Spydarr Summary to Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
