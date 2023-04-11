import os
import pandas as pd
import re
from datetime import datetime

from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest, Dimension, Metric, DateRange, OrderBy

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = 'XXXX'
property_id = 'XXXX'

client = BetaAnalyticsDataClient()

def run_report(start_date,end_date):
    start_date = start_date
    end_date = end_date
    client = BetaAnalyticsDataClient()
    request = RunReportRequest(
        property='properties/' + property_id,
        dimensions=[Dimension(name="date")],
        metrics=[Metric(name="totalUsers"),
                 Metric(name="Conversions"),  # Actualiza la métrica de transacciones aquí
                 Metric(name="EventValue")],
        order_bys=[OrderBy(dimension={'dimension_name': 'date'})],
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
    )
    response = client.run_report(request)
    return response

def export_to_csv(response):
    dimension_headers = [header.name for header in response.dimension_headers]
    metric_headers = [header.name for header in response.metric_headers]
    data = []
    for row in response.rows:
        row_data = {"date": None, "totalUsers": 0, "Conversions": 0, "EventValue": 0}
        for dimension, value in zip(dimension_headers, row.dimension_values):
            row_data[dimension] = value.value
        for metric, value in zip(metric_headers, row.metric_values):
            row_data[metric] = value.value
        data.append(row_data)
    df = pd.DataFrame(data)
    # Convertir los valores a los tipos de datos correctos
    df['totalUsers'] = df['totalUsers'].astype(int)
    df['Conversions'] = df['Conversions'].astype(int)
    df['EventValue'] = df['EventValue'].astype(float)
    # Cambiar cabeceras
    column_mapping = {
        "date": "Date",
        "totalUsers": "Total Users",
        "Conversions": "Conversions",
        "EventValue": "Conversions value",
    }
    df.rename(columns=column_mapping, inplace=True)
    dfs_by_month = {}
    for index, row in df.iterrows():
        date = row["Date"]
        month = date[4:6]
        year = date[0:4]
        month_name = None
        if month == "01":
            month_name = f"January {year}"
        elif month == "02":
            month_name = f"February {year}"
        elif month == "03":
            month_name = f"March {year}"
        elif month == "04":
            month_name = f"April {year}"
        elif month == "05":
            month_name = f"May {year}"
        elif month == "06":
            month_name = f"June {year}"
        elif month == "07":
            month_name = f"July {year}"
        elif month == "08":
            month_name = f"August {year}"
        elif month == "09":
            month_name = f"September {year}"
        elif month == "10":
            month_name = f"October {year}"
        elif month == "11":
            month_name = f"November {year}"
        elif month == "12":
            month_name = f"December {year}"    
        else:
            continue
        if month_name not in dfs_by_month:
            dfs_by_month[month_name] = []
        dfs_by_month[month_name].append(row)
    with pd.ExcelWriter("analytics_data.xlsx") as writer:
        for month_name, rows in dfs_by_month.items():
            month_df = pd.DataFrame(rows)
            month_df.to_excel(writer, sheet_name=month_name, index=False)

if __name__ == "__main__":
    print("Welcome, this program will export the data from Google Analytics via API to excel format.")
    def this_month():
        today = datetime.now()
        first_day_of_month = today.replace(day=1)
        today = today.strftime("%Y-%m-%d")
        first_day_of_month = first_day_of_month.strftime("%Y-%m-%d")
        print(f"Start date: ", first_day_of_month, "End date: ", today)
        report_response = run_report(first_day_of_month,today)
        export_to_csv(report_response)
        print("Data export as analytics_data.xlsx")
    def this_year():
        today = datetime.now()
        first_day_of_year = today.replace(day=1, month=1)
        today = today.strftime("%Y-%m-%d")
        first_day_of_year = first_day_of_year.strftime("%Y-%m-%d")
        print(f"Start date: ", first_day_of_year, "End date: ", today)
        report_response = run_report(first_day_of_year,today)
        export_to_csv(report_response)
        print("Data export as analytics_data.xlsx")
    def custom_date():
        def is_valid_date(date_str):
            pattern = r"\d{4}-\d{2}-\d{2}"
            if not re.fullmatch(pattern, date_str):
                return False
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                return False
            return True
        def is_start_date_before_end_date(start_date, end_date):
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
            return start_date_obj <= end_date_obj
        start_date = input("Type the start date, example: 2022-03-01\nDate: ")
        while not is_valid_date(start_date):
            print("Invalid start date format. Please enter the date in the format 'YYYY-MM-DD'.")
            start_date = input("Type the start date, example: 2022-03-01\nDate: ")
        end_date = input("Type the end date, example: 2022-03-01\nDate: ")
        while not is_valid_date(end_date) or not is_start_date_before_end_date(start_date, end_date):
            if not is_valid_date(end_date):
                print("Invalid end date format. Please enter the date in the format 'YYYY-MM-DD'.")
            else:
                print("End date should be after or on the start date.")
            end_date = input("Type the end date, example: 2022-03-01\nDate: ")
        print(f"Start date: {start_date}, End date: {end_date}")
        report_response = run_report(start_date,end_date)
        export_to_csv(report_response)
        print("Data export as analytics_data.xlsx")
    def switch(option):
        switcher = {
            1: this_month,
            2: this_year,
            3: custom_date
        }
        func = switcher.get(option, lambda: print("Invalid option"))
        func()
    print("Please, select an option:")
    print("1. This month")
    print("2. This year")
    print("3. Custom date")
    option = int(input("Enter the option number: "))
    switch(option)