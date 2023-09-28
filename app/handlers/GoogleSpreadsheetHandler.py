import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from typing import List



class GoogleSpreadsheetHandler:
    def __init__(self, credentials_file: str = "../credentials/google_sheet.json", spreadsheet_name: str = "Paveiktais darbs") -> None:
        self.credentials_file: str = credentials_file
        self.spreadsheet_name: str = spreadsheet_name
        self.client: gspread.Client = None
        self.spreadsheet: gspread.Spreadsheet = None

    def authenticate(self) -> None:
        scope: List[str] = ['https://spreadsheets.google.com/feeds',
                            'https://www.googleapis.com/auth/drive']
        creds: ServiceAccountCredentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.credentials_file, scope)
        self.client: gspread.Client = gspread.authorize(creds)

    def open_spreadsheet(self) -> None:
        if self.client is None:
            self.authenticate()

        self.spreadsheet: gspread.Spreadsheet = self.client.open(
            self.spreadsheet_name)

    def read_worksheet(self, worksheet_name: str) -> List[List[str]]:
        if self.spreadsheet is None:
            self.open_spreadsheet()

        worksheet: gspread.Worksheet = self.spreadsheet.worksheet(
            worksheet_name)
        return worksheet.get_all_values()

    def read_data_as_dataframe(self, worksheet_name: str = "Veidlapu atbildes: 1", columns: List[str] = None) -> pd.DataFrame:
        data: List[List[str]] = self.read_worksheet(worksheet_name)
        if columns is None:
            columns: List[str] = data[0]
            data: List[List[str]] = data[1:]
        df: pd.DataFrame = pd.DataFrame(data, columns=columns)
        return df

    def write_to_google_sheet(self, values: List[List[str]], sheet_name: str) -> None:
        """
        Writes the specified values to a Google Sheet.
        :param values: A list of lists containing the values to write to the Google Sheet.
        :param sheet_name: The name of the sheet to write the values to.
        """
        if self.spreadsheet is None:
            self.open_spreadsheet()

        # Access the specific sheet by name
        sheet: gspread.Worksheet = self.spreadsheet.worksheet(sheet_name)

        # Clear existing values
        sheet.clear()

        # Update the sheet with the new values in bulk
        # The 'A1' indicates the starting cell, and the update method will
        # expand as needed based on the size of 'values'
        cell_list: List[gspread.Cell] = sheet.range(
            'A1:' + gspread.utils.rowcol_to_a1(len(values), len(values[0])))

        for cell, value in zip(cell_list, (item for sublist in values for item in sublist)):
            cell.value = value

        sheet.update_cells(cell_list)


def main():
    reader = GoogleSpreadsheetHandler()
    data = reader.get_data_as_dataframe()


if __name__ == "__main__":
    main()
