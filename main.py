import os
import pandas as pd
from datetime import datetime

def convert_to_sql_format(date_str):
    date_formats = ["%d %b %Y %H:%M:%S", "%d %b %Y %H:%M:%S:%f", "%d-%b-%y", "%d %b %Y"]
    for fmt in date_formats:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return date_str

def process_dates(df, columns_to_convert):
    for column in columns_to_convert:
        df[column] = df[column].apply(lambda x: convert_to_sql_format(str(x)) if pd.notna(x) else x)
    return df

def main(input_file, output_file, columns_to_convert):
    if os.path.exists(output_file):
        os.remove(output_file)
    
    df = pd.read_excel(input_file)

    primary_column = columns_to_convert[0]
    df = df[df[primary_column].notna()]
    df = process_dates(df, columns_to_convert)

    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    input_file = 'schedule_date.xlsx'
    output_file = 'formatted_schedules.xlsx'
    columns_to_convert = ['schedule_date1', 'schedule_date2', 'insert_date', 'modified_date']

    main(input_file, output_file, columns_to_convert)