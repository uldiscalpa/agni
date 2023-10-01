import pandas as pd
import numpy as np
from typing import List

from handlers.ExcelHandler import ExcelHandler
from .ReportProcessor import ReportProcessor


class ReportGenerator:
    def __init__(self, input_file_path: str, output_file_path: str):
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path

    def generate_report(self, report_type: str, **kwargs) -> List:
        # Generate the report from the DataFrame
        report_processor = ReportProcessor(
            self.input_file_path, self.output_file_path)
        report = report_processor.generate_basic_report(report_type, **kwargs)
        return report

    def write_report_to_excel(self, report: List, sheet_name: str) -> None:
        # Write the report back to the Excel file
        print(report)
        ExcelHandler.write_values_to_excel_staticmethod(file_path=self.output_file_path,
                                                        values=report, sheet_name=sheet_name)

    def run(self, project_type, **kwargs) -> None:
        # Generate the report

        if project_type == 'employees':
            report = self.generate_report(project_type, **kwargs)
            # Write the report back to the Excel file
            self.write_report_to_excel(report, sheet_name='Paveiktais darbs')

        elif project_type == 'project':
            report = self.generate_report(project_type)
            # Write the report back to the Excel file
            self.write_report_to_excel(report, sheet_name='Projekti')
