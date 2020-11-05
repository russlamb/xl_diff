"""
This module is intended to summarize the output of a comparison.  It can be called independently of the comparison
module on completed excel comparisons

here is the usage generated from the argparse module

usage: summary.py [-h] [--output_path OUTPUT_PATH] [--open OPEN] input_path

Summarize comparison output and create an excel file with results

positional arguments:
  input_path            Path to input file. This file should be an excel sheet that contains a comparison between two other workbooks. E.g. the output from compare.py

optional arguments:
  -h, --help            show this help message and exit
  --output_path OUTPUT_PATH, -o OUTPUT_PATH
                        Path to output file. If file exists it will be overwritten. The result will contain a spreadsheet with a list of columns with differences in the input sheet
  --open OPEN, -p OPEN  if true, open output file on completion using os.system. Output file path must resolve to a file. Adds quotes around file name so that paths with spaces can resolveon windows machines.



"""
import argparse
import os
import logging

from xl_diff import write_summary_file

logging.basicConfig(level=logging.DEBUG, format='%(asctime)-15s %(message)s')


# named tuple with class wrapped around for a docstring


def configure_arg_parser():
    """
    instantiates and configures an ArgumentParser class object.  Each argument is a mandatory or optional parameter
    that can be invoked from the command line.
    :return: argument parser object
    """
    parser = argparse.ArgumentParser(description="Summarize comparison output and create an excel file with results")
    parser.add_argument("input_path", help="Path to input file.  This file should be an excel sheet that contains"
                                           + " a comparison between two other workbooks. E.g. the output from compare.py"
                        )
    parser.add_argument("--output_path", '-o',
                        help="Path to output file.  If file exists it will be overwritten.  The result will contain"
                             + " a spreadsheet with a list of columns and number of differences."
                             + " if no output file is specified, new file name will be the same as input file "
                             + "with '_summary' appended")
    parser.add_argument("--open", '-p', type=bool, default=True, help="if true, open output file on completion " +
                                                                      "using os.system.  Output file path must " +
                                                                      "resolve to a file.  Adds quotes around file " +
                                                                      "name so that paths with spaces can resolve" +
                                                                      "on windows machines.")  # open on finish
    return parser


if __name__ == "__main__":
    parser = configure_arg_parser()
    args = parser.parse_args()
    in_path = args.input_path
    if not args.output_path:  # if no output path specified, use the input path and append _summary before extension
        file_parts = os.path.splitext(in_path)
        out_path = file_parts[0] + "_summary" + file_parts[1]
    else:
        out_path = args.output_path

    logging.info(f"input path {in_path}")
    logging.info(f"output path {out_path}")

    write_summary_file(in_path, out_path)
    if args.open:
        os.system(out_path)
