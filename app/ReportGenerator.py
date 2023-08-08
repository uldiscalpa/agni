
import pandas as pd
from datetime import date

from XLSXDataWriter import XLSXDataWriter
from GoogleSpreadsheetReader import GoogleSpreadsheetReader


class StoneStockReportGenerator:
    def __init__(self, incoming_file, outgoing_file, template_file):
        self.incoming_file = incoming_file
        self.outgoing_file = outgoing_file
        self.template_file = template_file

    def read_data(self, file_path):
        # Read data from Excel file using pandas
        data = pd.read_excel(file_path)
        return data

    def generate_report(self, outputfile_name):
        # Read incoming and outgoing data
        incoming_data = self.read_data(self.incoming_file)
        outgoing_data = self.read_data(self.outgoing_file)
        outgoing_data.rename(columns={
                             'Plāksnes numurs': 'Nr.'}, inplace=True)

        outgoing_total = outgoing_data.groupby('Nr.').agg({
            'Izgrieztie kvadrāti': 'sum',
            'Bilde ar izgriezto': lambda x: ' ; '.join(x.astype(str)),
            'Projekts': lambda x: ' ; '.join(x.astype(str))
        }).reset_index()

        # Merge incoming and outgoing data using 'ProductID' as the key and 'outer' merge
        merged_data = pd.merge(
            incoming_data, outgoing_total, on='Nr.', how='left')

        # Calculate current balance by subtracting outgoing quantity from incoming quantity
        merged_data['Atlikums'] = merged_data['Atl. Sāk.'] - \
            merged_data['Izgrieztie kvadrāti']

        merged_data.drop('Izlietoti m2', axis='columns', inplace=True)
        merged_data.rename(
            columns={'Izgrieztie kvadrāti': 'Izlietoti m2'}, inplace=True)
        merged_data.fillna(0, inplace=True)
        new_data = []
        new_data.append(merged_data.columns.tolist())
        new_data.extend(merged_data.values.tolist())

        writer = XLSXDataWriter(
            '../templates/stone_stock.xlsx', data_row_index=7, header_row_index=6)
        writer.append_data(new_data)
        # writer.apply_style(len(new_data)-1)
        writer.save(outputfile_name)


def main_DataFilesUpdater():

    credentials_file = "../credentials/google_sheet.json"
    spreadsheet_name = "stone_reports"
    worksheet_name = "answers"  # Replace with the name of your actual worksheet

    reader = GoogleSpreadsheetReader(credentials_file, spreadsheet_name)
    data = reader.read_worksheet(worksheet_name)
    # print(data)

    template_file = "../data/akmens_izgriezumi.xlsx"
    header_row_index = 1
    # data_row_index = 7
    # column_mapping = {"Plāksnes numurs": "Kods"}

    writer = XLSXDataWriter(template_file, header_row_index)
    writer.append_data(data)


def main():

    # Convert relative path to absolute path

    incoming_file = "C:/Users/User/Documents/Projects/Agni/data/akmens_plaksnes.xlsx"
    outgoing_file = "C:/Users/User/Documents/Projects/Agni/data/akmens_izgriezumi.xlsx"
    template_file = "C:/Users/User/Documents/Projects/Agni/templates/stone_stock.xlsx"
    output_file_name = '../' + str(date.today()) + ' stone.xlsx'

    report_generator = StoneStockReportGenerator(
        incoming_file, outgoing_file, template_file)
    report_generator.generate_report(output_file_name)


if __name__ == "__main__":
    main()
