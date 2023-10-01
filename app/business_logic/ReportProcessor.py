
import pandas as pd
from datetime import date
from typing import List, Dict


EXCEL_DATE_COLUMN_NAME = 'Formatēts datums'
EXCEL_EMPLOYEE_NAME_COLUMN_NAME = 'Vārds'
EXCEL_CONTRACT_PATH = '../data/employees.xlsx'


class ReportProcessor:
    INCOMING_FILE_COLUMN_NAMES = {
        'Vārds': 'employee',
        'Datums': 'date',
        'Laika darbs': 'time_work',
        'Gabala darbs': 'piece_work',
        'Daudzums gabala darbam': 'piece_work_quantity',
        'Patērētais laiks': 'time_consumed',
        'Samaksa': 'payment',
        'Projekts': 'project',
        'Darbs': 'work',
        'Laika zīmogs': 'time_stamp'
    }

    OUTGOING_FILE_COLUMN_NAMES = {
        'employee': 'Vārds',
        'date': 'Datums',
        'time_work': 'Laika darbs',
        'piece_work': 'Gabala darbs',
        'piece_work_quantity': 'Daudzums gabala darbam',
        'time_consumed': 'Patērētais laiks',
        'payment': 'Samaksa',
        'project': 'Projekts',
        'work': 'Darbs',
        'time_stamp': 'Laika zīmogs'
    }

    def __init__(self, incoming_file: str, outgoing_file: str) -> None:
        self.incoming_file = incoming_file
        self.outgoing_file = outgoing_file

    def read_data(self, file_path: str) -> pd.DataFrame:
        # Read data from Excel file using pandas
        data = pd.read_excel(file_path)
        return data

    def filter_data_by_time_span(self, data: pd.DataFrame, start_date: str, end_date: str) -> pd.DataFrame:
        filtered_data = data[(data['Datums'].dt.date >= start_date) & (
            data['Datums'].dt.date <= end_date)]
        return filtered_data

    def generate_basic_report(self, report_type: str, *args: List, **kwargs: Dict) -> pd.DataFrame:
        if report_type == 'employees':
            return self.generate_employee_report(*args, **kwargs)
        elif report_type == 'project':
            return self.generate_project_report(*args, **kwargs)
        else:
            raise ValueError('Invalid report type')

    def generate_project_report(self) -> List[List]:
        # Read incoming and outgoing data
        incoming_data = self.read_data(self.incoming_file)
        outgoing_data = self.read_data(self.outgoing_file)
        outgoing_data.rename(columns={'Plāksnes numurs': 'Nr.'}, inplace=True)

        # Calculate total outgoing quantity by 'Nr.'
        outgoing_total = outgoing_data.groupby('Nr.').agg({
            'Izgrieztie kvadrāti': 'sum',
            'Bilde ar izgriezto': lambda x: ' ; '.join(x.astype(str)),
            'Projekts': lambda x: ' ; '.join(x.astype(str))
        }).reset_index()

        # Merge incoming and outgoing data using 'Nr.' as the key and 'outer' merge
        merged_data = pd.merge(
            incoming_data, outgoing_total, on='Nr.', how='left')

        # Calculate current balance by subtracting outgoing quantity from incoming quantity
        merged_data['Atlikums'] = merged_data['Atl. Sāk.'] - \
            merged_data['Izgrieztie kvadrāti']

        # Clean up merged data
        merged_data.drop('Izlietoti m2', axis='columns', inplace=True)
        merged_data.rename(
            columns={'Izgrieztie kvadrāti': 'Izlietoti m2'}, inplace=True)
        merged_data.fillna(0, inplace=True)

        # Convert merged data to list of lists
        new_data = []
        new_data.append(merged_data.columns.tolist())
        new_data.extend(merged_data.values.tolist())

        return new_data

    def generate_employee_report(self, start_date: str, end_date: str, employee_list: List[str], *args, **kwargs) -> List[List]:
        # Read contract data
        df_contract = pd.read_excel(
            EXCEL_CONTRACT_PATH, sheet_name='Contracts')

        # clean data
        df = self.clean_employee_report_data()

        
        #  calcluate salary
        df = df.merge(df_contract, on=['Darbs', 'Vārds'], how='left',).reset_index()
        df.reset_index(drop=True, inplace=True)
        df['Darbs'].fillna('default')

        df['Kopā'] = df['Daudzums'] * df['Samaksa']
        df['Kopā'] = df['Kopā'].round(2)

        df = df.reset_index(drop=True)
        filtered_data = self.filter_data_by_time_span(
            df, start_date, end_date)
        # Convert the filtered data to a list of values with the first row as the column list
        column_list = filtered_data.columns.tolist()
        data_list = filtered_data.values.tolist()
        data_list.insert(0, column_list)
        return data_list

    def clean_employee_report_data(self) -> pd.DataFrame:
        # Clean up source data
        df = self.read_data(self.incoming_file)
        #  clean date data
        df.drop(['Laika zīmogs', 'Datums'], axis=1, inplace=True)
        df['Datums'] = pd.to_datetime(
            df['Formatēts datums'], format='%d.%m.%Y')
        df.drop(['Formatēts datums'], axis=1, inplace=True)

        #  clean job list data
        df['Darbs'] = df['Laika darbs'].combine_first(
            df['Gabala darbs'])
        df['Darbs'].fillna('default', inplace=True)
        df.drop(['Laika darbs', 'Gabala darbs'], axis=1, inplace=True)

        # clean amount data
        df['Daudzums gabala darbam'] = df['Daudzums gabala darbam'].str.replace(
            ',', '.')  # replace comma with dot

        df['Patērētais laiks'] = df['Patērētais laiks'].str.replace(
            ',', '.')  # replace comma with dot

        df['Daudzums'] = df['Patērētais laiks'].combine_first(
            df['Daudzums gabala darbam']).astype(float)
        df['Daudzums'].fillna(0, inplace=True)
        df.drop(['Patērētais laiks', 'Daudzums gabala darbam'],
                axis=1, inplace=True)

        return df

    def generate_pivot_table(self, data: List[List], group_by, values, index, aggfunc) -> List[List]:
        df = pd.DataFrame(data[1:], columns=data[0])
        pivot_table = pd.pivot_table(
            df, index=index, columns=group_by, values=values, aggfunc=aggfunc)
        pivot_table.reset_index(inplace=True)
        pivot_table.columns = pivot_table.columns.map(''.join)
        pivot_table.fillna(0, inplace=True)
        new_data = []
        new_data.append(pivot_table.columns.tolist())
        new_data.extend(pivot_table.values.tolist())
        return new_data


def main():
    pass


if __name__ == "__main__":
    main()
