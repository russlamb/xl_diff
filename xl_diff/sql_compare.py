"""
This module contains logic for comparing output from two sql queries run on two database connections

2020-09-10  RL  Created

"""
import logging
from concurrent.futures._base import as_completed
from concurrent.futures.thread import ThreadPoolExecutor

from .sql_to_xl import SqlToXl
from .compare import compare_files


class SqlCompare():
    """
    This class can be used to run sql on two data connections
    """

    def __init__(self, left_connection_string, right_connection_string, left_file_path=None,
                 right_file_path=None, left_sheet=None, right_sheet=None, multithreaded=False):
        """
        :param left_connection_string: connection string for left data connection (pyodbc)
        :param right_connection_string: connection string for right data connection (pyodbc)
        :param left_file_path: file path to store the results from the left connection
        :param right_file_path: file path to store results from right connection
        :param left_sheet: name of sheet in left output file
        :param right_sheet: name of sheet in right output file
        :param multithreaded: if True, run left and right sql simultaneously
        """
        self.left_connection_string = left_connection_string
        self.right_connection_string = right_connection_string
        self.left_file_path = r".\left.xlsx" if not left_file_path else left_file_path  # set default value
        self.right_file_path = r".\right.xlsx" if not right_file_path else right_file_path  # set default value
        self.left_sheet = "Sheet1" if not left_sheet else left_sheet  # set default value
        self.right_sheet = "Sheet1" if not right_sheet else right_sheet  # set default value
        self.multi_threaded = multithreaded  # if True, use threading to run left and right query simultaneously

    def generate_files_multithreaded(self, query, query_right=None):
        query_to_run_left = query
        query_to_run_right = query
        if query_right:  # if only one query is supplied, run the same query on both connections
            query_to_run_right = query_right  # if a second query is supplied for the right side, set it here.

        left_stx = SqlToXl(self.left_connection_string)
        right_stx = SqlToXl(self.right_connection_string)

        futures = []
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures.append(
                executor.submit(left_stx.save_sql, *[query_to_run_left, self.left_file_path, self.left_sheet]))
            futures.append(
                executor.submit(right_stx.save_sql, *[query_to_run_right, self.right_file_path, self.right_sheet]))

        for f in as_completed(futures):
            if f.exception():
                logging.error("recived Exception from thread {}".format(f.exception()))
                raise f.exception()
            else:
                logging.info("recived result from thread {}".format(f.result()))

        return self.left_file_path, self.right_file_path

    def generate_files_from_query(self, query, query_right=None):
        """
        Generate XLSX files from query from both left and right data connections
        :param query: SQL query to run on left connection.  This query is also run on right connection if a query_right is not supplied
        :param query_right: SQL query to run on right connection
        :return: tuple with left file path and right file path
        """
        if self.multi_threaded:  # if multithreading, use the multithreaded function
            return self.generate_files_multithreaded(query, query_right)

        query_to_run = query
        logging.info("Running SQL on left connection")
        left_sx = SqlToXl(self.left_connection_string)

        left_sx.save_sql(query, self.left_file_path, self.left_sheet)

        if query_right:  # if only one query is supplied, run the same query on both connections
            query_to_run = query_right  # if a second query is supplied for the right side, set it here.

        logging.info("Running SQL on right connection")
        right_sx = SqlToXl(self.right_connection_string)
        right_sx.save_sql(query_to_run, self.right_file_path, self.right_sheet)

        logging.info("Finished running SQL on both connections")
        return self.left_file_path, self.right_file_path

    def compare_query_results(self, output_path, query, query_right=None, threshold=0.001, open_on_finish=False,
                              sort_column=None, compare_type="default", has_header=True, sheet_matching="name",
                              add_summary=True):
        """
        Generate XLSX files from query on left and right connections, then compare the results.
        :param add_summary: if true, add a sheet with count of differences by column
        :param threshold: maximum acceptable differrnces of numerical values
        :param open_on_finish: if true, the output file will be opened when it is complete
        :param sort_column: numerical index of column, or list of such indices, used to sort rows
        :param compare_type: sorted or default (unsorted)
        :param has_header: if true, first row is excluded from sort
        :param sheet_matching: if "name", sheets with the same name will be compared.  Otherwise, order is used.
        :param output_path: path to output file
        :param query: SQL query for left connection.  Also used for right connection if query_right is not supplied
        :param query_right: SQL for right connection (optional)
        :return: path to output file
        """
        left_path, right_path = self.generate_files_from_query(query, query_right)
        compare_files(left_path, right_path, output_path, threshold, open_on_finish, sort_column, compare_type,
                      has_header, sheet_matching, add_summary)
        return output_path


def run_sql_comparison(left_connection_string, right_connection_string, output_path, query, query_right=None,
                       left_file_path=None, right_file_path=None, threshold=0.001, open_on_finish=False,
                       sort_column=None, compare_type="default", has_header=True, sheet_matching="name",
                       add_summary=True, multithreaded=False):
    """
    Instantiate SqlCompare and call to the compare function
    :param left_connection_string: connection string for left data connection (pyodbc)
    :param right_connection_string: connection string for right data connection (pyodbc)
    :param left_file_path: file path to store the results from the left connection
    :param right_file_path: file path to store results from right connection
    :param add_summary: if true, add a sheet with count of differences by column
    :param threshold: maximum acceptable differrnces of numerical values
    :param open_on_finish: if true, the output file will be opened when it is complete
    :param sort_column: numerical index of column, or list of such indices, used to sort rows
    :param compare_type: sorted or default (unsorted)
    :param has_header: if true, first row is excluded from sort
    :param sheet_matching: if "name", sheets with the same name will be compared.  Otherwise, order is used.
    :param output_path: path to output file
    :param query: SQL query for left connection.  Also used for right connection if query_right is not supplied
    :param query_right: SQL for right connection (optional)
    :param multithreaded: If True, run queries in parallel
    :return:
    """
    sc = SqlCompare(left_connection_string, right_connection_string, left_file_path, right_file_path,
                    multithreaded=multithreaded)
    return sc.compare_query_results(output_path, query, query_right, threshold, open_on_finish, sort_column,
                                    compare_type, has_header, sheet_matching, add_summary)
