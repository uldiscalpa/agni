import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from datetime import date
from ReportGenerator import StoneStockReportGenerator
from tkinter import messagebox
from ttkthemes import ThemedStyle


class StoneStockReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Stone Stock Report Generator")

        style = ThemedStyle(self)
        style.set_theme("arc")  # Set the 'radiance' theme
        self.set_window_icon()
        self.create_widgets()

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

    def browse_file(self, entry):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")])
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def browse_incoming_file(self):
        self.browse_file(self.incoming_entry)

    def browse_outgoing_file(self):
        self.browse_file(self.outgoing_entry)

    def browse_template_file(self):
        self.browse_file(self.template_entry)

    def generate_report(self):
        incoming_file = self.incoming_entry.get()
        outgoing_file = self.outgoing_entry.get()
        template_file = self.template_entry.get()
        output_file_name = '../' + str(date.today()) + ' stone.xlsx'

        report_generator = StoneStockReportGenerator(
            incoming_file, outgoing_file, template_file)
        report_generator.generate_report(output_file_name)
        messagebox.showinfo("File Generation Complete",
                            f"The report {output_file_name} has been generated successfully!")


if __name__ == "__main__":
    app = StoneStockReportApp()
    app.mainloop()
