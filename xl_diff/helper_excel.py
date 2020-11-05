"""
This module contains common helper functions to perform excel related tasks.
"""
import openpyxl


def get_empty_workbook():
    """
    Create workbook object and remove default worksheet
    :return: Workbook object
    """
    # get empty workbook
    output_wb = openpyxl.Workbook()
    output_sheets = output_wb.sheetnames
    for sheet in output_sheets:
        output_wb.remove(output_wb[sheet])
    return output_wb