import pandas as pd


class ExcelReader:
    def __init__(self, file_path):
        self.file_path = file_path

    def read_column_values(self, column_names: list[str]):
        """Reads the Excel file and extracts the values from the specified column."""
        # Read the Excel file into a DataFrame
        df = pd.read_excel(self.file_path)

        # Get the values from the specified column and sort them in descending order
        print(*column_names)
        column_values = df[column_names].drop_duplicates().sort_values(
            by=column_names[0], ascending=False).values.tolist()

        return column_values


def main():
    file_path = "..\\data\\projects.xlsx"
    column_name = "code"
    excel_reader = ExcelReader(file_path, column_name)
    values = excel_reader.read_column_values()
    print(values)


if __name__ == "__main__":
    main()
