"""
This module contains a class used to execute sql code and write the results to an excel XLSX file
"""
import logging
import pyodbc
import datetime

from .helper_excel import get_empty_workbook


class SqlToXl():
    """Use to connect to database, run sql, write results to Excel file"""

    def __init__(self, connection_string):
        """
        initialize the object by storing the connection string
        :param connection_string:connection string
        """
        self.connection_string = connection_string

    def save_sql(self, sql, filename, sheetname="Sheet1"):
        """
        Run the SQL on the specified connection and save the results in Excel.  File name and sheet names can be
        specified as arguments
        :param sql: SQL to run on the target database
        :param filename: Target Excel file name
        :param sheetname: Target sheet name in Excel file
        :return: None
        """
        print(self.connection_string)
        with pyodbc.connect(self.connection_string) as cnxn:
            rowid = 1
            colid = 1
            try:
                sql_to_run = sql
                logging.info(f"run sql: {sql_to_run}")  # log sql used

                cursor = cnxn.cursor()
                cursor.execute(sql_to_run)

                if not cursor.description:  # if cursor description is None, query might be using a stored procedure
                    noCount = """ SET NOCOUNT ON; """  # add SET NOCOUNT statement to sql executed
                    sql_to_run = noCount + sql  # this will correct the issue pyodbc has with stored procedures
                    cursor.execute(sql_to_run)  # rerun the sql with modification

                wb = get_empty_workbook()  # had issues using active sheet, so instead we remove the default sheet
                ws = wb.create_sheet("query_result")  # create new sheet
                ws.title = sheetname

                try:  # wrapping this in a try catch in case there are more issues
                    columns = [column[0] for column in cursor.description]
                except Exception as ex:
                    logging.error("Couldn't get column names from cursor description.  Check the query syntax"
                                  + " You may need to add SET NOCOUNT ON to your query "
                                  + f" Error info {ex}")
                    raise ex # reraise the exception

                for col in columns:
                    ws.cell(row=rowid, column=colid).value = col
                    colid += 1

                for row in cursor:
                    rowid += 1
                    colid = 1
                    for col in row:
                        ws.cell(row=rowid, column=colid).value = col
                        colid += 1

                logging.info(f"Saving {filename}")
                wb.save(filename)
                logging.info("Saved {} @ {}".format(filename, datetime.datetime.now()))
            except Exception as e:
                logging.error(f"An error occurred at while saving query results to executable. Error info {e}, "
                              + f"Line of output: {rowid}, Column index {colid}")
                raise (e)  # reraise error

    def get_query_results(self, query):
        """
        This can be used for testing connectivity to a database amd checking if data exists
        :param query: SQL query to run
        :return: list of lists.  Each inner list is a row returned from the SQL query
        """
        lines = []
        with pyodbc.connect(self.connection_string) as cnxn:
            cursor = cnxn.cursor()
            cursor.execute(query)

            for row in cursor:
                line = []
                for col in row:
                    line.append(col)
                lines.append(line)
        return lines


