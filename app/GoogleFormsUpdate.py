import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd


CREDENTIALS = "../credentials/google_sheet.json"
EXCEL_FILE_PATH = "../data/projects.xlsx"
GOOGLE_SPREADSHEET_NAME = "Test form (Responses)"

GOOGLE_SHEET_NAME = "Projekti"


class ExcelToGoogleSheetWithPandas:
    def __init__(self):
        # Set up Google Sheets credentials
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            CREDENTIALS, self.scope)
        self.client = gspread.authorize(self.creds)

    def read_excel_with_pandas(self, excel_file):
        # Read Excel file into a DataFrame
        df = pd.read_excel(excel_file)

        # Sort the DataFrame by the 'code' column
        df_sorted = df.sort_values(by='code', ascending=True)

        # Convert the sorted DataFrame to a list of lists
        values = df_sorted.values.tolist()

        values.insert(0, df.columns.tolist())

        return values

    def write_to_google_sheet(self, values, google_sheet_name, sheet_name):
        # Open the Google Spreadsheet
        spreadsheet = self.client.open(google_sheet_name)

        # Access the specific sheet by name
        sheet = spreadsheet.worksheet(sheet_name)

        # Clear existing values
        sheet.clear()

        # Update the sheet with the new values in bulk
        # The 'A1' indicates the starting cell, and the update method will
        # expand as needed based on the size of 'values'
        cell_list = sheet.range(
            'A1:' + gspread.utils.rowcol_to_a1(len(values), len(values[0])))

        for cell, value in zip(cell_list, (item for sublist in values for item in sublist)):
            cell.value = value

        sheet.update_cells(cell_list)

    def transfer_from_excel_to_gsheet(self, excel_file=EXCEL_FILE_PATH, google_sheet_name=GOOGLE_SPREADSHEET_NAME, sheet_name=GOOGLE_SHEET_NAME):
        values = self.read_excel_with_pandas(excel_file)
        self.write_to_google_sheet(values, google_sheet_name, sheet_name)


def main():
    handler = ExcelToGoogleSheetWithPandas()
    handler.transfer_from_excel_to_gsheet()


if __name__ == "__main__":
    main()
