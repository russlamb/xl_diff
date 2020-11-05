"""
This module will process multiple sql comparisons by parsing a file of arguments.  For each line in the file,
a separate comparison will be run
"""
import csv
import logging
from distutils.util import strtobool

from .sql_compare import run_sql_comparison
from .compare import compare_files


def process_file(has_header_flag, input_file, multithreaded=False, compare_only=False):
    count = 0
    with open(input_file) as f:
        s = csv.reader(f, delimiter='\t')
        for row in s:
            count += 1
            if has_header_flag and count == 1:
                continue
            else:
                logging.info(row)  # log the parameters

                # parse row into variables
                (left_connection_string, right_connection_string, output_path, query, query_right, left_file,
                 right_file, threshold, open_on_finish, sort_column, compare_type, has_header, sheet_matching,
                 add_summary, *_) = row  # underscore with star to capture extra columns

                sort_column_list = [int(x) for x in sort_column.split(',')]  # parse sort columns into list object
                logging.info("sort columns {}".format(sort_column_list))

                # store variables as a tuple
                parsed_values = (left_connection_string, right_connection_string, output_path, query, query_right,
                                 left_file, right_file, float(threshold), bool(strtobool(open_on_finish)),
                                 sort_column_list, compare_type, bool(strtobool(has_header)),
                                 bool(strtobool(sheet_matching)), bool(strtobool(add_summary)))

                logging.info("parsed values: {}".format(parsed_values))  # log tuple
                # run the comparison
                if compare_only:
                    compare_parameters = (left_file, right_file, output_path, float(threshold),
                                          bool(strtobool(open_on_finish)),sort_column_list,compare_type,
                                          bool(strtobool(has_header)), bool(strtobool(sheet_matching)),
                                          bool(strtobool(add_summary)))
                    logging.info("Compare parameters {}".format(compare_parameters))
                    compare_files(*compare_parameters)
                else:
                    run_sql_comparison(*parsed_values, multithreaded=multithreaded)  # unpack tuple as arguments
