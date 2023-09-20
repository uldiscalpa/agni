import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from ttkthemes import ThemedStyle
from tkcalendar import DateEntry
from datetime import datetime, timedelta

from ExcelReader import ExcelReader
from ExcelReport import ExcelReport
from GoogleSpreadsheetReader import GoogleSpreadsheetReader
from GoogleFormsUpdate import ExcelToGoogleSheetWithPandas
from ExcelUpdater import ExcelUpdater

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

    def create_widgets(self):
        ttk.Label(self, text="Incoming File:").grid(
            row=0, column=0, padx=5, pady=5)
        self.incoming_entry = ttk.Entry(self, width=50)
        self.incoming_entry.grid(row=0, column=1, padx=5, pady=5)
        self.incoming_entry.insert(
            0, "../data/akmens_plaksnes.xlsx")  # Predefined value

        incoming_browse_button = ttk.Button(
            self, text="Browse", command=self.browse_incoming_file)
        incoming_browse_button.grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(self, text="Outgoing File:").grid(
            row=1, column=0, padx=5, pady=5)
        self.outgoing_entry = ttk.Entry(self, width=50)
        self.outgoing_entry.grid(row=1, column=1, padx=5, pady=5)
        self.outgoing_entry.insert(
            0, "../data/akmens_izgriezumi.xlsx")  # Predefined value

        outgoing_browse_button = ttk.Button(
            self, text="Browse", command=self.browse_outgoing_file)
        outgoing_browse_button.grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(self, text="Template File:").grid(
            row=2, column=0, padx=5, pady=5)
        self.template_entry = ttk.Entry(self, width=50)
        self.template_entry.grid(row=2, column=1, padx=5, pady=5)
        self.template_entry.insert(
            0, "../templates/stone_stock.xlsx")  # Predefined value

        template_browse_button = ttk.Button(
            self, text="Browse", command=self.browse_template_file)
        template_browse_button.grid(row=2, column=2, padx=5, pady=5)

        generate_button = ttk.Button(
            self, text="Generate Report", command=self.generate_report)
        generate_button.grid(row=3, column=0, columnspan=3, padx=5, pady=10)


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


class BaseReportFields(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

    def read_column_data(self, file_path, column_list):
        project_list = ExcelReader(file_path=file_path).read_column_values(
            column_names=column_list)
        if len(column_list) > 1:
            return [f"{a} - {b}" for a, b in project_list]
        return project_list


class EmployeeFields(BaseReportFields):
    def __init__(self, parent):
        super().__init__(parent)

        # Calculate the first and last day of the previous month
        self.employees = super().read_column_data(
            file_path=EMPLOYEES_FILE_PATH, column_list=EMPLOYEES_COLUMN_NAME)

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
        report = ExcelReport(report_type="employees", save_path="..\data\path_to_save.xlsx",
                             source_data="..\data\paveiktais_darbs.xlsx")
        report.generate(**values)
        messagebox.showinfo("Success", "Report is created")
        # except Exception as e:
        #     # If there's an error during report generation
        #     print(e)
        #     messagebox.showerror("Error", f"An error occurred: {e}")


class ProjectFields(BaseReportFields):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(padx=10, pady=10)
        self.project_list = super().read_column_data(
            file_path=PROJECT_FILE_PATH, column_list=['code', 'name'])
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
            value = self.get_field_values()
            report = ExcelReport(report_type="employees", save_path="path_to_save.xlsx",
                                 source_data=f"..\data\{datetime.today().strftime('%Y%M %H:%M:%S')}_darbinieki.xlsx")
            report.generate(**value)
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
        # You can instantiate the class and call its method here
        excel_to_gsheet = ExcelToGoogleSheetWithPandas()
        excel_to_gsheet.transfer_from_excel_to_gsheet()

    def pull_data(self):
        # Similarly, instantiate the class and call its method here
        gspread_reader = GoogleSpreadsheetReader()
        df = gspread_reader.get_data_as_dataframe()
        ExcelUpdater('..\data\paveiktais_darbs.xlsx').combine_and_update(df)

    def save_config(self):
        # Just a dummy function to simulate saving the configuration.
        # In a real application, this would involve storing the configuration to a file or database.
        print(f"Saved configuration: {self.option_var.get()}")
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
