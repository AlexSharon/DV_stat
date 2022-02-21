from datetime import date as date_i

import pandas as pd
import os
import sqlite3 as sql

conn = sql.connect('CEAC.db')

df = pd.read_csv('FY2022-ceac.csv')
df = df.loc[(df['region'] == 'EU')]
df_clean = df.loc[df['status'] != 'None']

for i in range(df_clean.shape[0]):
    for column in (5, 6):
        date = df_clean.iat[i, column]
        if date != 'None':
            date = date.replace("Jan", "01")
            date = date.replace("Feb", "02")
            date = date.replace("Mar", "03")
            date = date.replace("Apr", "04")
            date = date.replace("May", "05")
            date = date.replace("Jun", "06")
            date = date.replace("Jul", "07")
            date = date.replace("Aug", "08")
            date = date.replace("Sep", "09")
            date = date.replace("Oct", "10")
            date = date.replace("Nov", "11")
            date = date.replace("Dec", "12")
            year = date[6:]
            month = date[3:5]
            day = date[0:2]
            date = year + '-' + month + '-' + day
            df_clean.iat[i, column] = date


df_clean = df_clean.drop(columns=['caseNumberFull'])
df_clean['caseNumber'] = pd.to_numeric(df_clean['caseNumber'])
df_clean.to_sql('ceac', conn, if_exists="replace", chunksize=500)

df_output = pd.DataFrame(columns=['Case_range', 'Issued', 'AP', 'Refused', 'Transfer_in_Progress', 'Ready',
                                  'In_Transit', 'AT_NVC', 'TOTAL', 'NVC_share'])

tracer = 0
i = 0
for cn in range(1000, 29000, 1000):
    df_chunk = df_clean.loc[(df_clean['caseNumber'] <= cn) & (df_clean['caseNumber'] > tracer)]
    tracer = cn

    issued = df_chunk.loc[df_chunk['status'] == 'Issued'].shape[0]
    ap = df_chunk.loc[df_chunk['status'] == 'Administrative Processing'].shape[0]
    refused = df_chunk.loc[df_chunk['status'] == 'Refused'].shape[0]
    transfer = df_chunk.loc[df_chunk['status'] == 'Transfer in Progress'].shape[0]
    ready = df_chunk.loc[df_chunk['status'] == 'Ready'].shape[0]
    transit = df_chunk.loc[df_chunk['status'] == 'In Transit'].shape[0]
    nvc = df_chunk.loc[df_chunk['status'] == 'At NVC'].shape[0]
    total = df_chunk.shape[0]
    nvc_share = nvc / float(total)

    df_output.loc[i] = [cn,
                        issued,
                        ap,
                        refused,
                        transfer,
                        ready,
                        transit,
                        nvc,
                        total,
                        nvc_share]
    i += 1

df_output.to_sql('stat', conn, if_exists="replace", chunksize=500)
file_name = str(date_i.today()) + "_report.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

df_for_report = [df_clean, df_output]
sheet_names = ['ceac', 'stat']
for df, name in zip(df_for_report, sheet_names):
    #df = df.sort_values('caseNumber')

    df.to_excel(writer, sheet_name=name, startrow=1, header=False, index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[name]

    # Get the dimensions of the dataframe.
    (max_row, max_col) = df.shape

    # Create a list of column headers, to use in add_table().
    column_settings = [{'header': column} for column in df.columns]

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    # Add some cell formats.
    format1 = workbook.add_format({'num_format': '### ##0.00'})
    format2 = workbook.add_format({'num_format': '0.00%'})

    # Set the column width and format.
    if name == 'ceac':
        worksheet.set_column('A:A', 5, None)  # region
        worksheet.set_column('B:B', 9, None)  # caseNumber
        worksheet.set_column('C:C', 8, None)  # consulate
        worksheet.set_column('D:D', 10, None)  # status
        worksheet.set_column('E:E', 11, None)  # submitDate
        worksheet.set_column('F:F', 11, None)  # statusDate
        worksheet.set_column('G:G', 8, None)  # issued
        worksheet.set_column('H:H', 8, None)  # AP
        worksheet.set_column('I:I', 8, None)  # Ready
        worksheet.set_column('J:J', 8, None)  # Refused
        worksheet.set_column('K:K', 8, None)  # in transit
        worksheet.set_column('L:L', 8, None)  # transfer
        worksheet.set_column('M:M', 8, None)  # NVC
        worksheet.set_column('N:N', 8, None)  # potAP
    elif name == 'stat':
        worksheet.set_column('A:A', 10, None)  # case
        worksheet.set_column('B:B', 10, None)  # issued
        worksheet.set_column('C:C', 10, None)  # ap
        worksheet.set_column('D:D', 10, None)  # refused
        worksheet.set_column('E:E', 10, None)  # transfer
        worksheet.set_column('F:F', 10, None)  # ready
        worksheet.set_column('G:G', 10, None)  # transit
        worksheet.set_column('H:H', 8, None)  # NVC
        worksheet.set_column('I:I', 9, None)  # total
        worksheet.set_column('J:J', 9, format2)  # percent
    else:
        pass
writer.save()
