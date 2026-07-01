import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series

def generate_report_with_native_chart():
    # 1. Load data from the master transaction sheet
    file_path = "IT_Procurement_Project.xlsx"
    df_po = pd.read_excel(file_path, sheet_name="Purchase_Orders")
    df_sup = pd.read_excel(file_path, sheet_name="Suppliers")
    df_po['PO_Date'] = pd.to_datetime(df_po['PO_Date'])

    # 2. Compute David Simchi-Levi metrics
    # Calculate delivery unpredictability (Standard Deviation)
    lt_stats = df_po.groupby('Supplier_ID')['Lead_Time'].agg(['mean', 'std']).reset_index()
    lt_stats.columns = ['Supplier_ID', 'Average_Lead_Time', 'Lead_Time_Variability_Sigma']
    lt_stats['Lead_Time_Variability_Sigma'] = lt_stats['Lead_Time_Variability_Sigma'].fillna(0)

    # Calculate financial pipeline velocity per supplier
    total_days = (df_po['PO_Date'].max() - df_po['PO_Date'].min()).days
    supplier_spend = df_po.groupby('Supplier_ID')['Total_Amount'].sum().reset_index()
    supplier_spend['Daily_Spend_Velocity'] = supplier_spend['Total_Amount'] / total_days

    # Merge into risk register master frame
    df_analysis = df_sup.merge(lt_stats, on='Supplier_ID', how='inner')
    df_analysis = df_analysis.merge(supplier_spend[['Supplier_ID', 'Daily_Spend_Velocity']], on='Supplier_ID', how='inner')
    df_analysis['Simulated_TTR_Days'] = np.round(df_analysis['Risk_Score'] * 12) 
    df_analysis['Risk_Exposure_Index_REI'] = df_analysis['Daily_Spend_Velocity'] * df_analysis['Simulated_TTR_Days']
    df_analysis = df_analysis.sort_values(by='Risk_Exposure_Index_REI', ascending=False).reset_index(drop=True)

    # 3. Initialize OpenPyXL Workbook Structure
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Risk_Register"
    ws.views.sheetView[0].showGridLines = True  # Ensure gridlines remain visible

    # Add data layout
    headers = ['Supplier_ID', 'Supplier_Name', 'Category', 'Country', 'Risk_Score', 
               'Average_Lead_Time', 'Lead_Time_Variability_Sigma', 'Daily_Spend_Velocity', 
               'Simulated_TTR_Days', 'Risk_Exposure_Index_REI']
    ws.append(headers)

    # Style header row with classic corporate Navy theme
    navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    white_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = navy_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center")

    # Inject metric rows
    for _, row in df_analysis.iterrows():
        ws.append([
            row['Supplier_ID'], row['Supplier_Name'], row['Category'], row['Country'], row['Risk_Score'],
            row['Average_Lead_Time'], row['Lead_Time_Variability_Sigma'], row['Daily_Spend_Velocity'],
            row['Simulated_TTR_Days'], row['Risk_Exposure_Index_REI']
        ])

    # Apply precise cell financial and decimal formatting
    for row_idx in range(2, len(df_analysis) + 2):
        ws.cell(row=row_idx, column=6).number_format = "0.0"     # Avg Lead Time
        ws.cell(row=row_idx, column=7).number_format = "0.00"    # Lead Time Sigma
        ws.cell(row=row_idx, column=8).number_format = "$#,##0"  # Spend Velocity
        ws.cell(row=row_idx, column=10).number_format = "$#,##0" # REI

    # 4. Construct Native Interactive Scatter Chart Component
    chart = ScatterChart()
    chart.title = "Simchi-Levi Procurement Risk Matrix"
    chart.style = 13
    chart.x_axis.title = 'Lead-Time Variability (Sigma - Unpredictability)'
    chart.y_axis.title = 'Risk Exposure Index (REI - Total Recovery Liability)'

    # Define chart references: X = Column 7 (Sigma), Y = Column 10 (REI)
    xvalues = Reference(ws, min_col=7, min_row=2, max_row=len(df_analysis) + 1)
    yvalues = Reference(ws, min_col=10, min_row=2, max_row=len(df_analysis) + 1)
    series = Series(yvalues, xvalues, title_from_data=False)
    
    # Configure clean data points without ugly connective lines
    series.marker.symbol = "circle"
    series.marker.size = 7
    series.graphicalProperties.line.noFill = True 

    chart.series.append(series)
    chart.width = 17
    chart.height = 10
    
    # Render the chart at position cell L2
    ws.add_chart(chart, "L2")

    # 5. Calculate and apply optimal column scaling widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Save output report
    wb.save("Simchi_Levi_Risk_Report_With_Chart.xlsx")
    print("Workbook successfully saved with embedded native chart elements.")

if __name__ == "__main__":
    generate_report_with_native_chart()# Simchi-Levi-Matrix-Lead-time-Variability-Vs-Systemic-Risk-Exposure
