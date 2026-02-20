import pandas as pd

# Detailed Data for Thursday (extracted from the provided text and image)
# CE Data: Strike, OI, % Change
thursday_ce_data = [
    [26000, 148.1, 0.38], [25800, 147.0, 1.37], [25700, 123.1, 1.72],
    [25600, 114.0, 4.00], [25900, 94.5, 0.95], [25500, 76.7, 1.50],
    [26400, 71.1, 0.94], [26200, 69.7, 0.54], [26100, 63.7, 0.53],
    [26300, 62.7, 0.56], [25750, 62.5, 1.41], [25850, 62.2, 1.63],
    [25650, 41.0, 4.36], [25950, 38.2, 0.71], [26250, 37.0, 0.70]
]

# PE Data: Strike, OI, % Change
thursday_pe_data = [
    [25000, 102.8, 0.20], [24500, 81.5, 0.08], [24800, 62.6, 0.82],
    [25400, 60.6, 0.17], [25500, 56.3, -0.38], [25200, 47.4, 0.00],
    [25100, 43.5, 0.40], [25300, 40.9, 0.02], [25600, 38.3, -0.44],
    [24700, 38.1, -0.02], [25700, 33.3, -0.63], [26000, 31.1, 0.03],
    [24900, 30.9, 0.16], [25800, 29.8, -0.45], [25350, 21.5, -0.09]
]

# Create DataFrame for detailed Thursday data
detailed_df = pd.DataFrame({
    'Strike_CE': [x[0] for x in thursday_ce_data],
    'CE_OI': [x[1] for x in thursday_ce_data],
    '%OI_CE': [x[2] for x in thursday_ce_data],
    'Strike_PE': [x[0] for x in thursday_pe_data],
    'PE_OI': [x[1] for x in thursday_pe_data],
    '%OI_PE': [x[2] for x in thursday_pe_data]
})

# Dynamically calculate Maximum OI for the summary
max_ce_thurs = detailed_df['CE_OI'].max()
max_pe_thurs = detailed_df['PE_OI'].max()

# Summary Data for the Week (Monday to Friday)
# Based on the user's data, only Thursday has valid OI entries.
summary_data = {
    'Day': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
    'Max_CE_OI': [0.0, 0.0, 0.0, max_ce_thurs, 0.0],
    'Max_PE_OI': [0.0, 0.0, 0.0, max_pe_thurs, 0.0]
}
summary_df = pd.DataFrame(summary_data)

output_file = 'OI_Max_Report.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Write summary table
    summary_df.to_excel(writer, sheet_name='Max OI Summary', index=False)
    # Write detailed table
    detailed_df.to_excel(writer, sheet_name='Thursday Details', index=False)

    workbook  = writer.book
    summary_ws = writer.sheets['Max OI Summary']

    # --- Add Chart to Summary Sheet ---
    chart = workbook.add_chart({'type': 'column'})

    # Configure CE Series
    chart.add_series({
        'name':       'Max CE OI',
        'categories': "='Max OI Summary'!$A$2:$A$6",
        'values':     "='Max OI Summary'!$B$2:$B$6",
        'fill':       {'color': '#4472C4'} # Blue
    })

    # Configure PE Series
    chart.add_series({
        'name':       'Max PE OI',
        'categories': "='Max OI Summary'!$A$2:$A$6",
        'values':     "='Max OI Summary'!$C$2:$C$6",
        'fill':       {'color': '#ED7D31'} # Orange
    })

    # Chart Styling
    chart.set_title({'name': 'Weekly Maximum Open Interest (CE vs PE)'})
    chart.set_x_axis({'name': 'Day of the Week'})
    chart.set_y_axis({'name': 'Max Open Interest'})
    chart.set_legend({'position': 'bottom'})

    # Insert chart into the summary worksheet
    summary_ws.insert_chart('E2', chart)

    # --- Formatting for Thursday Details Sheet ---
    details_ws = writer.sheets['Thursday Details']
    percent_fmt = workbook.add_format({'num_format': '0.0%'})

    # Set column widths and percentage format
    details_ws.set_column('A:A', 10)
    details_ws.set_column('B:B', 10)
    details_ws.set_column('C:C', 12, percent_fmt)
    details_ws.set_column('D:D', 10)
    details_ws.set_column('E:E', 10)
    details_ws.set_column('F:F', 12, percent_fmt)

    # Header formatting for both sheets
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9D9D9',
        'border': 1,
        'align': 'center'
    })

    for col_num, value in enumerate(summary_df.columns.values):
        summary_ws.write(0, col_num, value, header_format)
    for col_num, value in enumerate(detailed_df.columns.values):
        details_ws.write(0, col_num, value, header_format)

print(f"Successfully generated {output_file}")
