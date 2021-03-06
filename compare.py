"""
This module is the main entry point of the compare excel project.  This project compares two excel or CSV files
and saves the results in an excel file.

Command line parsing has been moved to a separate file to keep code clean.

To run the module, you need to pass in three filenames: the two files being compared, denoted "left" and "right",
and the output file.  In addition to the filenames, there are several optional arguments to pass in that will alter
the comparison behavior.  By default the comparison will be a cell-by-cell comparison.  Other options include
Sorting based on a single column (e.g. a primary key) or group of columns (e.g. a composite key)

The following usage information is generated by the argparse module and displayed to the console when running
the command "python compare.py --help":

usage: compare.py [-h] [--threshold THRESHOLD] [--open OPEN] [--compare_type COMPARE_TYPE] [--sort_column SORT_COLUMN | --sort_column_list SORT_COLUMN_LIST [SORT_COLUMN_LIST ...]] [--no_header] [--sheet_matching SHEET_MATCHING] [--convert_csv CONVERT_CSV]
                  left right output

Compare two Excel (XLSX) files sheet by sheet. Create a new excel file with values side by side along with differences.

positional arguments:
  left                  Path to first file for comparison. Can be CSV or XLSX. In output file, these values will be on left
  right                 Path to second file for comparison. Can be CSV or XLSX. In output file, these values will be on right
  output                Path to output file. If file exists it will be overwritten. If compare_type is 'sorted' then it will contain copies of data from original files as well as the values side by side in a combined sheet.

optional arguments:
  -h, --help            show this help message and exit
  --threshold THRESHOLD, -t THRESHOLD
                        threshold for numeric values to be considered different. e.g. when threshold = 0.01 if left and right values are closer than 0,01 then consider the same. Mainly affects coloring of difference column for numeric values
  --open OPEN, -p OPEN  if true, open output file on completion using os.system. Output file path must resolve to a file. Adds quotes around file name so that paths with spaces can resolveon windows machines.
  --compare_type COMPARE_TYPE, -c COMPARE_TYPE
                        if set to 'sorted', the comparison tool will attempt to line up each side based on the values of sort_column specified. 'default' is a cell-by-cell comparison.
  --sort_column SORT_COLUMN, -s SORT_COLUMN
                        numeric offset (1-based) of column to use for sorting. To be used for a primary key. if compare type is 'sorted', this column will be used to sort and line up each sideE.g. 1 would be the first column.
  --sort_column_list SORT_COLUMN_LIST [SORT_COLUMN_LIST ...], -l SORT_COLUMN_LIST [SORT_COLUMN_LIST ...]
                        space separated list of numbers for columns to use for sorting. E.g. '-l 1 2' would be both the first and second columns. E.g. a compound key. if compare type is 'sorted', these columns will be used to sort and line up each side
  --no_header, -d       if sheets do not have headers, set this flag so headers can be excluded from comparison
  --sheet_matching SHEET_MATCHING, -m SHEET_MATCHING
                        can be either 'name' or 'order'. If name, only sheets with the same name are compared. if order, sheets are compared in order. E.g. 1st sheet vs 1st sheet.
  --convert_csv CONVERT_CSV, -v CONVERT_CSV
                        if True, convert csv files to xlsx

"""
import argparse
import logging

from xl_diff.validators import is_number
from xl_diff import compare_files

logging.basicConfig(level=logging.DEBUG, format='%(asctime)-15s %(message)s')


def run_from_command_line():
    """
    Parse command line arguments and run comparison of excel / CSV files
    :return: None
    """
    # Test command: python compare.py tests/left.xlsx tests/right.xlsx tests/output.xlsx -c sorted -l 1
    parser = compare_excel_configure_arg_parser()
    args = parser.parse_args()
    sort_column_arg = args.sort_column if args.sort_column else args.sort_column_list
    if args.compare_type == "sorted":
        if not is_number(args.sort_column) and not isinstance(args.sort_column_list,
                                                              list):  # sort key can be number or list
            parser.error("sort column must be a number, or sort column list must be a list if compare type is sorted")
    has_header_flag = True
    if args.no_header:  # check for header argument.  parameter for the compare function needs to be inverted
        has_header_flag = not args.no_header
    logging.info("Has Header: {}".format(has_header_flag))
    logging.info(f"sorting column: {sort_column_arg}")
    compare_files(args.left, args.right, args.output, args.threshold, args.open, sort_column_arg, args.compare_type,
                  has_header_flag, args.sheet_matching, args.summary)


def compare_excel_configure_arg_parser():
    """
    instantiates and configures an ArgumentParser class object.  Each argument is a mandatory or optional parameter
    that can be invoked from the command line.
    :return: argument parser object
    """
    parser = argparse.ArgumentParser(description="Compare two Excel (XLSX) files sheet by sheet.  " +
                                                 "Create a new excel file with values side by side along with " +
                                                 "differences.")  # inisial description
    parser.add_argument("left", help="Path to first file for comparison.  Can be CSV or XLSX.  In output file, these " +
                                     "values will be on left")  # first workbook file

    parser.add_argument("right", help="Path to second file for comparison.  Can be CSV or XLSX.  In output file, " +
                                      " these values will be on right")  # second workbook file
    parser.add_argument("output", help="Path to output file.  If file exists it will be overwritten.  " +
                                       "It will contain copies of data from original " +
                                       "files as well as the values side by side in a combined sheet.")  # output file
    parser.add_argument("--threshold", '-t', type=float, default=0.001,
                        help="threshold for numeric values to be considered different.  e.g. when threshold = 0.01 " +
                             "if left and right values are closer than 0,01 then consider the same.  Mainly affects " +
                             "coloring of difference column for numeric values")  # max allowed difference
    parser.add_argument("--open", '-p', type=bool, default=True, help="if true, open output file on completion " +
                                                                      "using os.system.  Output file path must " +
                                                                      "resolve to a file.  Adds quotes around file " +
                                                                      "name so that paths with spaces can resolve" +
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
    parser.add_argument("--sheet_matching", "-m", default="order", help="can be either 'name' or 'order'.  If name, " +
                                                                        "only sheets with the same name are " +
                                                                        "compared. " +
                                                                        "if order, sheets are compared in order. " +
                                                                        "E.g. 1st sheet vs 1st sheet.")
    parser.add_argument("--convert_csv", "-v", default=True, help="if True, convert csv files to xlsx")
    parser.add_argument("--summary", "-y", type=bool, default=True, help="if True, add a summary sheet to the output."
                                                                         "Summary module assumes every 3rd sheet in "
                                                                         "workbook contains comparison.")

    return parser


if __name__ == "__main__":
    run_from_command_line()
