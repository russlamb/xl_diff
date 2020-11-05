"""
This module contains logic for comparing output from two sql queries run on two database connections

This module is intended to be called as an entry point from the command line

2020-09-10  RL  Created

"""
import argparse

from xl_diff import run_sql_comparison
import logging

from xl_diff.validators import is_number

logging.basicConfig(level=logging.DEBUG, format='%(asctime)-15s %(message)s')


def run_from_command_line():
    # Test command: python sql_compare.py <left connection> <right connection> <output file> <query> -c sorted -l 1
    parser = sql_compare_configure_arg_parser()
    args = parser.parse_args()

    # parse sort column arg
    sort_column_arg = args.sort_column if args.sort_column else args.sort_column_list
    if args.compare_type == "sorted":  # sort key can be number or list
        if not is_number(args.sort_column) and not isinstance(args.sort_column_list, list):
            parser.error("sort column must be a number, or sort column list must be a list if compare type is sorted")

    # parse header arg
    has_header_flag = True
    if args.no_header:  # check for header argument.  parameter for the compare function needs to be inverted
        has_header_flag = not args.no_header

    logging.info("Has Header: {}".format(has_header_flag))
    logging.info(f"sorting column: {sort_column_arg}")

    # perform comparison
    run_sql_comparison(args.left, args.right, args.output, args.query, args.query_right, args.left_file, args.right_file
                       , args.threshold, args.open, sort_column_arg, args.compare_type, has_header_flag,
                       args.sheet_matching, args.summary, args.multithreaded)


def sql_compare_configure_arg_parser():
    """
    instantiates and configures an ArgumentParser class object.  Each argument is a mandatory or optional parameter
    that can be invoked from the command line.
    :return: argument parser object
    """
    parser = argparse.ArgumentParser(description="Compare query run on two data connections.  " +
                                                 "Create a new excel file with values side by side along with " +
                                                 "differences.")  # inisial description
    parser.add_argument("left", help="First Connection string.  In output file, these values will be on left")
    parser.add_argument("right", help="Second connection string.  In output file, these values will be on right")
    parser.add_argument("output", help="Path to output file.  If file exists it will be overwritten.  " +
                                       "It will contain copies of data from original " +
                                       "files as well as the values side by side in a combined sheet.")  # output file
    parser.add_argument("query", help="Query to run on left connection, or both connections if no query_right supplied")
    parser.add_argument("--query_right", '-Q', help="Query to run on right query, if different from left query")
    parser.add_argument("--left_file", '-L', help="File path to store results from left query.  left.xlsx by default")
    parser.add_argument("--right_file", '-R', help="File path to store results from right query. right.xlsx by default")
    parser.add_argument("--threshold", '-t', type=float, default=0.001,
                        help="threshold for numeric values to be considered different.  e.g. when threshold = 0.01 " +
                             "if left and right values are closer than 0,01 then consider the same.  Mainly affects " +
                             "coloring of difference column for numeric values")  # max allowed difference
    parser.add_argument("--open", '-p', type=bool, default=True,
                        help="if true, open output file on completion using os.system.  Output file path must " +
                             "resolve to a file.  Adds quotes around file name so that paths with spaces can resolve" +
                             "on windows machines.")  # open on finish
    parser.add_argument("--compare_type", '-c', default="default",  # sorted or cell-by-cell comparison
                        help="if set to 'sorted', the comparison tool will attempt to line up each side based on " +
                             "the values of sort_column specified.  'default' is a cell-by-cell comparison.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--sort_column", "-s", type=int, default=None,
                       help="numeric offset (1-based) of column to use for sorting.  " +
                            "To be used for a primary key. if compare type is 'sorted', this " +
                            "column will be used to sort and line up each side" +
                            "E.g. 1 would be the first column.")  # single sort column
    group.add_argument("--sort_column_list", "-l", nargs="+", type=int, default=None,  # multiple column sort
                       help="space separated list of numbers for columns to use for sorting.  E.g. '-l 1 2' would " +
                            "be both the first and second columns.  E.g. a compound key. if compare type is " +
                            "'sorted', these columns will be used to sort and line up each side")
    parser.add_argument("--no_header", "-d", action="store_true",  # inidicates first column is not a header
                        help="if sheets do not have headers, set this flag so headers can be excluded from comparison")
    parser.add_argument("--sheet_matching", "-m", default="order",
                        help="can be either 'name' or 'order'.  If name, only sheets with the same name are " +
                             "compared. if order, sheets are compared in order. E.g. 1st sheet vs 1st sheet.")
    parser.add_argument("--summary", "-y", type=bool, default=True,
                        help="if True, add a summary sheet to the output.  Summary module assumes every 3rd sheet in "
                             "workbook contains comparison.")
    parser.add_argument("--multithreaded", "-M", action="store_true",
                        help="run both sql commands at the same time using multithreading")

    return parser


if __name__ == "__main__":
    run_from_command_line()
