"""
This module contains a function used to convert a CSV file to an Excel XLSX workbook.
"""
import csv
import logging
import os

import openpyxl as xl


def convert_csv_to_excel(csv_path):
    """
    This function converts a csv file, given by its file path, to an excel file in the same directory with the same
    name.
    :param csv_path:string file path of CSV file to convert
    :return: string file path of converted Excel file.
    """
    (file_path, file_extension) = os.path.splitext(csv_path)  # split the csv pathname to remove the extension

    wb = xl.Workbook()  # create the excel workbook
    ws = wb.active  # use the active sheet by default
    logging.info("converting file to xlsx: '{}'".format(csv_path))

    with open(csv_path, newline='') as csv_file:  # append each row of the csv to the excel worksheet
        rd = csv.reader(csv_file, delimiter=",", quotechar='"')
        for row in rd:
            ws.append(row)

    output_path = os.path.join(file_path + '.xlsx')  # output file path should be the same as the csv file
    logging.info("saving to file: '{}'".format(output_path))
    wb.save(output_path)  # save the converted file
    return output_path
