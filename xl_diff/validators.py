"""
This module contains simple functions used to check the type of values or evaluate the validity of inputs
"""
import os

from dateutil.parser import parse

from .convert import convert_csv_to_excel


def is_number(my_value):
    """
    Check if value is a number or not by casting to a float and handling any errors.
    :param my_value: value to check
    :return: true if number, false if not.
    """
    try:
        float(my_value)
        return True
    except Exception:
        return False


def is_date(my_value):
    """
    Check if a value is a date by parsing it.  If an error occurs, it's not a date.
    :param my_value: value to check
    :return: true if date, false if not
    """
    try:
        parse(my_value)
        return True
    except Exception:
        return False


def is_extension(file_path, extension):
    """
    Check if the extension argument matches the extension in the file path
    :param file_path: file path to check
    :param extension: extension to compare
    :return: Boolean value.  True if match.
    """
    (file_name, file_extension) = os.path.splitext(file_path)
    return file_extension.lower() == extension


def is_file_extension_valid(file_path):
    """
    check if file has .csv file extension.  If file is invalid, throw an error.
    :param file_path: file to check
    :return: file path if valud.
    """
    if is_extension(file_path, '.csv'):
        file_path = convert_csv_to_excel(file_path)
    elif not is_extension(file_path, ".xlsx"):
        raise ValueError("file extension for {} is not xlsx or csv.  file cannot be processed.".format(file_path))
    return file_path
