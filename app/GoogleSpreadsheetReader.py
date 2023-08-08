import gspread
import openpyxl
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSpreadsheetReader:
    def __init__(self, credentials_file: str = "../credentials/google_sheet.json", spreadsheet_name: str = "stone_reports"):
        self.credentials_file = credentials_file
        self.spreadsheet_name = spreadsheet_name
        self.client = None
        self.spreadsheet = None

    def authenticate(self):
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            self.credentials_file, scope)
        self.client = gspread.authorize(creds)

    def open_spreadsheet(self):
        if self.client is None:
            self.authenticate()

        self.spreadsheet = self.client.open(self.spreadsheet_name)

    def read_worksheet(self, worksheet_name):
        if self.spreadsheet is None:
            self.open_spreadsheet()

        worksheet = self.spreadsheet.worksheet(worksheet_name)
        return worksheet.get_all_values()

    def read_and_remove_data(self, worksheet_name):
        data = self.read_worksheet(worksheet_name)

        if not data:
            return []  # No data in the worksheet

        num_rows_to_clear = len(data)
        print(len(data))
        if num_rows_to_clear > 0:
            self.spreadsheet.worksheet(
                worksheet_name).delete_rows(2, num_rows_to_clear)

        return data


# Example usage of the GoogleSpreadsheetReader class


def main():
    credentials_file = "../credentials/google_sheet.json"
    spreadsheet_name = "stone_reports"
    worksheet_name = "answers"  # Replace with the name of your actual worksheet

    reader = GoogleSpreadsheetReader(credentials_file, spreadsheet_name)
    data = reader.read_worksheet(worksheet_name)
    print(data)


if __name__ == "__main__":
    main()
