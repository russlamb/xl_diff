"""
This module will process multiple sql comparisons by parsing a file of arguments.  For each line in the file,
a separate comparison will be run
"""
import argparse
import logging
from xl_diff.sql_compare_file import process_file

logging.basicConfig(level=logging.DEBUG, format='%(asctime)-15s %(message)s')


def run_from_command_line():
    parser = sql_compare_file_configure_arg_parser()
    args = parser.parse_args()
    input_file = args.file

    # parse header parameter
    has_header_flag = True
    if args.no_header:  # check for header argument.  parameter for the compare function needs to be inverted
        has_header_flag = not args.no_header

    logging.info("Has Header: {}".format(has_header_flag))
    logging.info("Input file: {}".format(input_file))
    logging.info("Multithreading off: {}".format(args.multithreading_off))
    logging.info("Compare only: {}".format(args.compare_only))

    multithreaded = not args.multithreading_off #invert the boolean
    process_file(has_header_flag, input_file, multithreaded=multithreaded, compare_only=args.compare_only)


def sql_compare_file_configure_arg_parser():
    """
    instantiates and configures an ArgumentParser class object.  Each argument is a mandatory or optional parameter
    that can be invoked from the command line.



    :return: argument parser object
    """
    parser = argparse.ArgumentParser(description="Compare multiple queries based on file input.  " +
                                                 "Create a excel files with values side by side along with " +
                                                 "differences.", formatter_class=argparse.RawTextHelpFormatter)

    parser.add_argument("file", help="Path to Tab-delimited file containing list of arguments.  Each line will"
                                     + " invoke a new call of the sql_compare module.  "
                                     + """
    The order of arguments within the file are:
    Left Connection String  -   connection string
    Right Connection String -   connection string
    Output File             -   file path
    Query                   -   SQL query
    Query Right             -   SQL query (optional)
    Left File               -   file path
    Right File              -   file path
    Threshold               -   numeric value.  value differences below this threshold are considered identical
    Open on Finish          -   If true, open the output file when analysis is complete
    Sort Columns	        -   A comma-separated list of integers identifying columns that make up a Unique row
    Compare Type            -   "sorted" or "default".  when sorted, Sort Columns must be supplied.   
    Has Header              -   When true, skips the first row when sorting
    Sheet Matching          -   When true, attempts to match sheets by name.  Otherwise, uses order.
    Add Summary             -   When true, adds a summary page with count of differences by column
                        """)

    parser.add_argument("--no_header", "-d", action="store_true",  # inidicates first column is not a header
                        help="if flag is present, assumes sheets do not have headers.  Headers will be excluded from "
                             + "the list of commands")
    parser.add_argument("--multithreading_off", "-M", action="store_true",
                        help="if flag is present, do not run both sql commands at the same time via multithreading")
    parser.add_argument("--compare_only", "-C", action="store_true",
                        help="if flag is present, do not run sql.  Only perform the comparison.")

    return parser


if __name__ == "__main__":
    run_from_command_line()
