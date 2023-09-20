import pandas as pd
from configs import EMPLOYEES_SPREADSHEET_DATA_TYPES
from openpyxl import load_workbook
import numpy as np

EXCEL_DATE_COLUMN_NAME = 'Formatēts datums'
EXCEL_EMPLOYEE_NAME_COLUMN_NAME = 'Vārds'
EXCEL_CONTRACT_PATH = '../data/employees.xlsx'


class ExcelReport:
    def __init__(self, report_type, save_path, source_data):
        self.report_type = report_type
        self.save_path = save_path
        self.source_data = pd.read_excel(
            source_data)

    def generate_employee_report(self, start_date, end_date, *args, **kwargs):
        # Assuming the source data has a 'Date' column in datetime format and 'Employee' column
        df_contract = pd.read_excel(EXCEL_CONTRACT_PATH)
        df = self.source_data
        df.drop('Laika zīmogs', axis=1, inplace=True)
        df.drop('Datums', axis=1, inplace=True)


        # Merge on both 'TaskID' and 'Date'\

        df['Darbs'] = df['Laika darbs'].combine_first(
            df['Gabala darbs']).astype(str)
        df = df.merge(df_contract, on=['Darbs', 'Vārds'], how='left')

        # df['Patērētais laiks'] = df['Patērētais laiks'].str.replace(',', '.').astype(float)
        df['Darbs'].fillna('default')
        df['Daudzums'] = df['Patērētais laiks'].combine_first(
            df['Daudzums gabala darbam'])
        print(df['Daudzums'])
        # df['Daudzums'] = df['Daudzums'].str.replace(',', '.').astype(float)
        # df['Daudzums'] = df['Daudzums'].fillna(0)

        print(df['Daudzums'])
        df['Formatēts datums'] = pd.to_datetime(
            df['Formatēts datums'], format='%d.%m.%Y')
        df['Kopā'] = df['Daudzums'].astype(float) * df['Samaksa'].astype(float)

        print(df['Daudzums'])
        df.rename(columns={'Formatēts datums': 'Datums'}, inplace=True)
        filtered_data = df[
            (df['Datums'].dt.date >= start_date) &
            (df['Datums'].dt.date <= end_date)
            # (df[EXCEL_EMPLOYEE_NAME_COLUMN_NAME].isin(employee_list))
        ]
        print(filtered_data['Daudzums'])
        pivot = filtered_data.pivot_table(
            values=['Patērētais laiks'], index='Projekts', columns='Vārds', aggfunc='sum')
        pivot.loc['Total'] = pivot.sum()

        pivot_2 = filtered_data.pivot_table(
            values=['Kopā'], index=['Datums'], columns='Vārds', aggfunc='sum')
        pivot_2.loc['Total'] = pivot.sum()

        with pd.ExcelWriter(self.save_path) as writer:
            # Write each DataFrame to a specific sheet
            pivot.to_excel(writer, sheet_name='summary')
            pivot_2.to_excel(writer, sheet_name='salary')
            filtered_data.to_excel(writer, sheet_name='raw_data', index=False)

        adjust_columns_width(self.save_path)
        # Write to Excel
        # filtered_data.to_excel(self.save_path, sheet_name='raw_data', index=False)

    def generate(self, *args, **kwargs):
        print(kwargs.get('start_date'))
        if self.report_type == "employees":
            self.generate_employee_report(
                start_date=kwargs.get('start_date'),
                end_date=kwargs.get('end_date'),
            )
        # You can add other report types with additional elif conditions


def adjust_columns_width(file_name):
    # Load the previously saved workbook
    workbook = load_workbook(file_name)
    sheet = workbook.active

    # Loop through columns and set width
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]  # Convert tuple to list
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # adding a little extra space
        try:
            #  number 3 to escape merged cells
            sheet.column_dimensions[column[2].column_letter].width = adjusted_width
        except Exception as e:
            print(e)

    # Save the modified file
    workbook.save(file_name)


def main():
    # Assuming you have a DataFrame df with your data
    report = ExcelReport(report_type="employees",
                         save_path="path_to_save.xlsx", source_data=df)
    report.generate(start_date="2022-01-01",
                    end_date="2022-02-01", employee_list=['Alice', 'Bob'])


if __name__ == "__main__":
    main()
