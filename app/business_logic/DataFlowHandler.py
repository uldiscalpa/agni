import pandas as pd
from typing import List, Callable


class DataJoiner:
    def __init__(self, worksheet_name: str, excel_file_path: str, spreadsheet_handler: Callable, excel_handler: Callable) -> None:
        # Initialize the class with the name of the Google Spreadsheet, the name of the worksheet to read from, the path to the Excel file, the path to the credentials file, and optional handlers for the Google Spreadsheet and Excel file.
        self.worksheet_name = worksheet_name
        self.excel_file_path = excel_file_path
        self.google_spreadsheet_handler = spreadsheet_handler
        self.excel_handler = excel_handler

    def read_spreadsheet_data(self) -> pd.DataFrame:
        # Read data from Google Spreadsheet and return it as a pandas DataFrame.
        data = self.google_spreadsheet_handler.read_data_as_dataframe(
            self.worksheet_name)
        return data

    def read_excel_data(self) -> pd.DataFrame:
        # Read data from Excel file and return it as a pandas DataFrame.
        data = self.excel_handler.read_data()
        return data

    def join_data(self, spreadsheet_data: pd.DataFrame, excel_data: pd.DataFrame, on_columns: List[str] = []) -> pd.DataFrame:
        # Join the data from the Google Spreadsheet and the Excel file together and return it as a pandas DataFrame.
        joined_data = pd.concat([excel_data, spreadsheet_data], ignore_index=True).drop_duplicates(
            subset=on_columns, keep='last').reset_index(drop=True)
        return joined_data

    def write_data(self, data: pd.DataFrame) -> None:
        # Write data to Excel file.
        self.excel_handler.write_data_frame_to_excel(
            df=data, file_name=self.excel_file_path)

    def run(self) -> None:
        # Read data from Google Spreadsheet.
        spreadsheet_data = self.read_spreadsheet_data()
        # Read data from Excel file.
        excel_data = self.read_excel_data()
        # Join the data from the Google Spreadsheet and the Excel file together.
        joined_data = self.join_data(spreadsheet_data, excel_data, on_columns=[
                                     "Vārds", "Laika zīmogs"])
        # Write data to Excel file.
        self.write_data(joined_data)
