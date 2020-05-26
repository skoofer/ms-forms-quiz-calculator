import os
from sys import argv
import pandas as pd


def calculate_quiz_results(path, result_file_name):
    result_file_name_with_format = result_file_name + '.xlsx' if '.xlsx' not in result_file_name else result_file_name
    writer = pd.ExcelWriter(os.path.join(path, result_file_name_with_format), engine='xlsxwriter')

    df_total = pd.DataFrame()
    files = os.listdir(path)

    for file in (f for f in files if not f == result_file_name_with_format):
        category_name = file.replace('.xlsx', '')

        df = pd.read_excel(os.path.join(path, file), usecols='E,F')
        df = df.sort_values(by='Total points', ascending=False)
        df = df.reset_index(drop=True)
        df.index = df.index + 1

        df.to_excel(writer, sheet_name=category_name)

        worksheet = writer.sheets[category_name]
        worksheet.set_column(1, 2, 30)
        df_total = df_total.append(df)

    df_total = df_total.groupby(['Name']).sum()
    df_total = df_total.sort_values(by='Total points', ascending=False)
    df_total = df_total.reset_index()
    df_total.index = df_total.index + 1
    df_total.to_excel(writer, sheet_name='Grand Total')
    worksheet = writer.sheets['Grand Total']
    worksheet.set_column(1, 2, 30)

    writer.save()
    writer.close()

    print('Success!')


if __name__ == '__main__':
    path_arg = argv[1]
    result_file_name_arg = argv[2]

    calculate_quiz_results(path_arg, result_file_name_arg)
