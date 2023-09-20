import pandas as pd
from ExcelReport import adjust_columns_width 


class ExcelUpdater:

    def __init__(self, excel_path):
        self.excel_path = excel_path

    def fetch_local_data(self):
        return pd.read_excel(self.excel_path)

    def combine_and_update(self, google_data: pd.DataFrame):
        local_data = self.fetch_local_data()
        combined_data = pd.concat([local_data, google_data]).drop_duplicates(
            subset=['Laika zīmogs', 'Vārds'], keep='last')
        # combined_data.astype('datetime64[ns]')
        combined_data['Formatēts datums'] = pd.to_datetime(combined_data['Formatēts datums'], format='%d.%m.%Y')
        combined_data.sort_values(by='Formatēts datums', inplace=True)
        combined_data['Formatēts datums'] = combined_data['Formatēts datums'].dt.strftime('%d.%m.%Y')

        combined_data.to_excel(self.excel_path, index=False)
        adjust_columns_width(self.excel_path)


def main():
    pass

if __name__ == "__main__":
    main()