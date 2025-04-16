import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
df=pd.read_csv(r'teg.csv', index_col=0, sep=';',decimal=',')
print(df.info())

dft = df['Tag'].unique()

average_values = df.groupby('Tag').mean()
result_df = pd.DataFrame(average_values)
print(result_df)

# Save the average values to a single Excel file
excel_filename = 'average.xlsx'
writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
result_df.to_excel(writer, sheet_name='Average Values')
writer.close()
print(f"Excel file created: {excel_filename}")

# Create an Excel file for each tag and insert a chart
for tag in dft:
    dfft = df.loc[df['Tag'] == tag]
    dfft = dfft.sort_values(by='Date')

    # Create an Excel writer
    excel_filename = f'{tag}.xlsx'
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
    dfft.to_excel(writer, sheet_name='Потребление')

    # Access the workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Потребление']

    # Define the chart
    chart = workbook.add_chart({'type': 'line'})

    # Configure the series of the chart
    chart.add_series({
        'name': f'Tag {tag}',
        'categories': ['Потребление', 1, 0, len(dfft), 0],  # Date column
        'values': ['Потребление', 1, 2, len(dfft), 2],      # Value column
    })

    # Set the chart title and axes names
    chart.set_title({'name': f'График потребления {tag}'})
    chart.set_x_axis({'name': '', 'tick_font':{'name':'Times New Roman','size':10}})
    chart.set_y_axis({'name': 'P, МВт','tick_font':{'name':'Times New Roman','size':10}})

    # Remove the legend
    chart.set_legend({'none': True})

    # Insert the chart into the worksheet
    worksheet.insert_chart('G2', chart)

    # Save the Excel file
    writer.close()
    print(f"Excel file created: {excel_filename}")

print("Done!")
