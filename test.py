from datetime import date

import pandas as pd
import os

df = pd.read_csv ('FY2022-ceac.csv')
df = df.loc[(df['region'] == 'EU')]# & (df['status'] != 'None')]

#print(df)


file_name = str(date.today()) + "_report_all.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
df.to_excel(writer, sheet_name='invoices', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets["invoices"]
(max_row, max_col) = df.shape
column_settings = [{'header': column} for column in df.columns]
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

worksheet.set_column('A:A', 10, None)   # region
worksheet.set_column('B:B', 10, None)   # caseNumber
worksheet.set_column('C:C', 12, None)   # caseNumberFull
worksheet.set_column('D:D', 10, None)   # consulate
worksheet.set_column('E:E', 12, None)    # status
worksheet.set_column('F:F', 12, None)    #
worksheet.set_column('G:G', 12, None)    #
worksheet.set_column('H:H', 12, None)    #
worksheet.set_column('I:I', 12, None)    #
worksheet.set_column('J:J', 12, None)    #

writer.save()
