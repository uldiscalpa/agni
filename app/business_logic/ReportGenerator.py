import pandas as pd
import numpy as np
from typing import List

from handlers.ExcelHandler import ExcelHandler
from . import ReportProcessor


class ReportGenerator:
    def __init__(self, input_file_path: str, output_file_path: str):
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path

    def generate_report(self, report_type: str) -> List:
        # Generate the report from the DataFrame
        report_processor = ReportProcessor(self.input_file_path, self.output_file_path)
        report = report_processor.generate_report(report_type)
        return report

    def write_report_to_excel(self, report: List, sheet_name: str) -> None:
        # Write the report back to the Excel file
        excel_handler = ExcelHandler()
        excel_handler.write_values_to_excel_staticmethod(
            values=report, sheet_name=sheet_name, file_name=self.output_file_path)

    def run(self, project_type) -> None:
        # Generate the report
        
        if project_type == 'employees':
            report = self.generate_report(project_type)
             # Write the report back to the Excel file
            self.write_report_to_excel(report, sheet_name='Paveiktais darbs')

            pivot_report = self.generate_pivot_table(report, gorup_by='Vārds', values='Samaksa', index='Projekts', aggfunc=np.sum)
            self.write_report_to_excel(pivot_report, sheet_name='Kopā par projektiem')
        
        elif project_type == 'project':
            report = self.generate_report(project_type)
            # Write the report back to the Excel file
            self.write_report_to_excel(report, sheet_name='Projekti')



