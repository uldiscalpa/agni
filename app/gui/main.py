import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from ttkthemes import ThemedStyle
from tkcalendar import DateEntry
from datetime import datetime, timedelta

from business_logic.ReportGenerator import ReportGenerator
from business_logic.DataFlowHandler import DataJoiner
from handlers.ExcelHandler import ExcelHandler
from handlers.GoogleSpreadsheetHandler import GoogleSpreadsheetHandler

DATE_FORMAT = 'dd.mm.y'
PROJECT_FILE_PATH = '..\data\projects.xlsx'
PROJECT_COLUMN_NAME = ['code', 'name']

EMPLOYEES_FILE_PATH = '..\data\employees.xlsx'
EMPLOYEES_COLUMN_NAME = ['Vārds']


class AgniReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Stone Stock Report Generator")

        style = ThemedStyle(self)
        style.set_theme("arc")  # Set the 'radiance' theme
        self.set_window_icon()
        self.create_tabs()

    def create_tabs(self):
        notebook = ttk.Notebook(self)
        notebook.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

        # Tab 1 - Report Creation
        report_view = ReportCreationView(notebook)
        notebook.add(report_view, text="Report Creation")

        # Tab 2 - Configuration
        config_view = ConfigurationView(notebook)
        notebook.add(config_view, text="Configuration")

    def set_window_icon(self):
        try:
            # Replace 'icon.ico' with the path to your icon file
            self.iconbitmap("../assets/logo.png")
        except tk.TclError as e:
            # If setting the icon fails, ignore the error
            print(e)
            pass


class ReportCreationView(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        label = tk.Label(self, text="Atskaišu veidošana")
        label.pack(pady=20)

        # Dropdown menu for options

        self.dropdown_options = ["projects", "employees"]
        self.selected_option = tk.StringVar()
        self.selected_option.set(self.dropdown_options[1])  # Default option

        self.dropdown = ttk.Combobox(
            self, textvariable=self.selected_option, values=self.dropdown_options, width=50)
        self.dropdown.pack(pady=20)
        self.dropdown.bind("<<ComboboxSelected>>", self.update_fields)

        # Container frame for the fields that will change based on the dropdown selection
        self.fields_frame = tk.Frame(self)
        self.fields_frame.pack(pady=20, padx=20)

        self.update_fields()

    def update_fields(self, event=None):
        """Update fields based on dropdown selection."""
        for widget in self.fields_frame.winfo_children():
            widget.destroy()

        if self.selected_option.get() == "projects":
            self.project_fields = ProjectFields(self.fields_frame)
            self.project_fields.pack(fill=tk.BOTH, expand=True)
        elif self.selected_option.get() == "employees":
            self.employee_fields = EmployeeFields(self.fields_frame)
            self.employee_fields.pack(fill=tk.BOTH, expand=True)


class EmployeeFields(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # reading employees from excel file
        self.employees = ExcelHandler(
            "..\data\employees.xlsx").read_column_values(["Vārds"])
        self.create_widgets()

    def create_widgets(self):

        today = datetime.today()
        first_day_of_current_month = today.replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - \
            timedelta(days=1)
        first_day_of_previous_month = last_day_of_previous_month.replace(day=1)

        start_date_label = tk.Label(self, text="Start Date:")
        start_date_label.pack(anchor=tk.W, pady=(0, 5))

        self.start_date = DateEntry(self, date_pattern=DATE_FORMAT)
        self.start_date.set_date(first_day_of_previous_month)
        self.start_date.pack(fill=tk.X, pady=(0, 10))

        end_date_label = tk.Label(self, text="End Date:")
        end_date_label.pack(anchor=tk.W, pady=(0, 5))

        self.end_date = DateEntry(self, date_pattern=DATE_FORMAT)
        self.end_date.set_date(last_day_of_previous_month)
        self.end_date.pack(fill=tk.X)

        self.file_name = tk.Label(self, text="Faila nosaukums:")
        self.file_name.pack(pady=10)

        current_date = today.strftime('%Y-%m-%d')

        # Create an entry (text field)
        self.file_name = tk.Entry(self)
        self.file_name.insert(0, f"{current_date} darbinieki")
        self.file_name.pack(pady=10)

        self.generate_button = ttk.Button(
            self, text="Generate", command=self.generate_report)
        self.generate_button.pack(pady=20)

    def get_field_values(self):
        values = {}
        values.update({
            'start_date': self.start_date.get_date(),
            'end_date': self.end_date.get_date(),
            'employee_list': self.employees
        })
        return values

    def generate_report(self):
        values = self.get_field_values()
        # try:
        report = ReportGenerator(input_file_path=EMPLOYEES_FILE_PATH,
                                 output_file_path=f"..\reports\{self.file_name.get()}.xlsx")

        report.run(**values)
        messagebox.showinfo("Success", "Report is created")


class ProjectFields(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(padx=10, pady=10)

        # reading projects from excel file
        self.project_list = ExcelHandler(
            "..\data\projects.xlsx").read_column_values("code")
        self.create_widgets()

    def create_widgets(self):
        # Add a label
        self.label = ttk.Label(self, text="Select Projects:")
        self.label.pack(pady=5)  # Using pack here

        # Create a Listbox with the MULTIPLE select mode
        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE,  width=50)
        self.listbox.pack(pady=5)  # Using pack here

        # Sample project list (this can be loaded dynamically)

        # Add projects to the Listbox
        for project in self.project_list:
            self.listbox.insert(tk.END, project)

        # Button to get selected projects
        self.button = ttk.Button(
            self, text="Get Selected Projects", command=self.get_selected_projects)
        self.button.pack(pady=5)  # Using pack here

        self.generate_button = ttk.Button(
            self, text="Generate", command=self.generate_report)
        self.generate_button.pack(pady=20)

    def get_selected_projects(self):
        # Retrieve selected projects from the Listbox
        # This returns a tuple of selected indices
        selected_indices = self.listbox.curselection()
        selected_projects = [self.listbox.get(i) for i in selected_indices]

        # Print or return the selected projects (for demonstration, we're printing them)

    def get_field_values(self):
        values = super().get_field_values()
        selected_indices = self.listbox.curselection()
        # Retrieve base field values
        selected_projects = [self.listbox.get(i) for i in selected_indices]
        values.update({
            'selected_project': selected_projects
        })
        return values

    def generate_report(self):
        try:
            messagebox.showinfo("Success", "Report is created")
        except Exception as e:
            # If there's an error during report generation
            messagebox.showerror("Error", f"An error occurred: {e}")


class ConfigurationView(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def create_widgets(self):
        # "Update form" button
        self.update_form_button = ttk.Button(
            self, text="Update form", command=self.update_form)
        self.update_form_button.pack(pady=10)

        # "Pull data" button
        self.pull_data_button = ttk.Button(
            self, text="Pull data", command=self.pull_data)
        self.pull_data_button.pack(pady=10)

        self.save_button = ttk.Button(
            self, text="Save Configuration", command=self.save_config)
        self.save_button.pack(pady=10)

    def update_form(self):
        # Read data from Excel file
        excel_handler = ExcelHandler(file_path='..\data\projects.xlsx')
        data = excel_handler.read_column_values(['code', 'name'])
        data = [f'{row[0]} - {row[1]}' for row in data]
        # Write data to Google Spreadsheet
        spreadsheet_handler = GoogleSpreadsheetHandler(
            credentials_file='..\credentials\google_sheet.json', spreadsheet_name='Test form (Responses)')
        spreadsheet_handler.write_to_google_sheet(
            values=data, sheet_name='Projekti')

    def pull_data(self):
        # Read data from Google Spreadsheet and Excel file, join it, and write it to Excel file
        spreadsheet_handler = GoogleSpreadsheetHandler(
            credentials_file='..\credentials\google_sheet.json', spreadsheet_name='Paveiktais darbs')
        excel_handler = ExcelHandler(file_path='..\data\paveiktais_darbs.xlsx')
        data_joiner = DataJoiner(spreadsheet_handler=spreadsheet_handler, excel_handler=excel_handler,
                                 worksheet_name='Veidlapu atbildes: 1', excel_file_path='..\data\employees.xlsx')
        data_joiner.run()

    def save_config(self):
        # Just a dummy function to simulate saving the configuration.
        # In a real application, this would involve storing the configuration to a file or database.
        print("Configuration saved successfully!")
        messagebox.showinfo(
            "Configuration", "Configuration saved successfully!")


class DataMergingView(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        label = ttk.Label(self, text="Informācijas apmaiņa")
        label.pack(pady=20)

        change_button = ttk.Button(
            self, text="Uz sākumu", command=parent.show_first_view)
        change_button.pack(pady=20)


if __name__ == "__main__":
    app = AgniReportApp()
    app.mainloop()
