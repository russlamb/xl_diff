# Automatically Compare Two Excel, CSV Files, or SQL Queries 

Created On: November 20 2019 10:06 AM

## Overview 
This tool compares excel documents, CSV files, or SQL queries column by column and identifies differences.  

This was designed for testing the output of a report in two different environments, but can be used for any excel
 document compatible with openpyxl <sup>1.</sup>.  

Comparisons are done by identifying if values being compared are numbers, dates, or strings.  If numbers, the 
difference between the values are stored in a difference column.  If dates or strings, the word "Same" or "Different" 
is stored.  

The output file contains each column of the input files (named "left" and "right") alongside the difference column.
The script proceeds until it has compared every sheet in the file.  

The script can use the values of a single column to line up the rows of the input files, so that missing values do 
not offset the comparison.

If the input files are CSV, they are converted to XLSX before the comparison begins via the excel comparison module.

When performing a SQL query comparison, SQL queries are run on two different database connections and saved as Excel 
before using the excel comparison module.   

<sup>1.</sup> Note: as of writing there are some limitations of openpyxl when working with Excel documents containing images.

### Key benefits include

1.  Output file contains values being compared side by side with differences, reducing context switching. This 
means better ease of testing and reduces human error.

1.  Sort by a unique row identifer (single column or multi-column) to Line up columns visually so missing values are 
identified quickly and do not interfere with the comparison

1.  Compare excel files and SQL queries quickly and automatically, meaning less developer time wasted moving data 
between Excel documents or doing comparisons


## Setup
There are several modules that can be called directly 

1. `compare.py` - compare excel files
2. `sql_compare.py` - compare two SQL queries run on two database connections
3. `sql_compare_file.py` - compare multiple SQL queries using a file input
4. `summary.py` - analyse an output file and add a summary file

Each module can be called with the `--help` argument to find out how to use it.  

### Prepare Python Environment

The necessary packages are found in the requirements.txt file. Install using pip and you should be good to go.

For example 
`pip install -r requirements.txt `

While not required, `pyinstaller` is recommended for building portable executable files for ease-of-use.   

#### Example
Here is an example of calling the `compare` module directly using the example data from the `tests` directory. 

`python compare.py tests\left.xlsx tests\right.xlsx tests\output.xlsx -s 1 -c sorted`

Here is an example of calling the tool using a compiled executable version with the same options applied.

`compare_excel.exe  tests\left.xlsx tests\right.xlsx tests\output.xlsx -s 1 -c sorted`

in this example, the first file to compare is the 'left.xlsx' file in the "tests" directory. 
the second file to compare is 'right.xlsx' in the tests directory.  
The target output file is 'output.xlsx' in the tests directory.  
The unique identifier column "-s" is in column 1, and the comparison type "-c" is 'sorted'.

For more details on usage, pass the `--help` argument to the module.

## Comparing Excel Files
For more information about parameters and options, pass the argument `--help` to the `compare` module.

### Basic usage
1.  Call the module by passing in 3 file paths for left input, right input, and output files.
    1.  if you want input files to be sorted and lined up by a specified column first, you must pass set 
    `--compare_type` or `-c` flag to `sorted`.  
    2.  You must also set `--sort_column` or `s` flag to a numeric value corresponding to the column you want to sort by 
(1-based, so first column = 1)
    3.  For multiple columns, use the `--sort_column_list` or `-l` flag and enter multiple numbers separated by 
    spaces.  Column order affects the sorting of the data.  E.g. `-l 3 1` will sort by third then first columns   
2.  If you want to open the output file automatically, set --open to True
3.  If your file does not have headers, pass the arguments --has_headers False

### Testing
Tests refer to sample XLSX and CSV files in the `tests` folder.

## Comparing SQL Results
For more information about parameters and options, pass the argument "--help" to the `sql_compare` module.

### Testing
Tests use a Sqlite 3 database in the `test_db` folder.  You can find the SQLite OBDC driver 
[here](http://www.ch-werner.de/sqliteodbc/)  

If you want to view the data, I recommend the utility [DB Browser for SQLite](https://sqlitebrowser.org/)

## Comparing multiple SQL results
For more information about parameters and options, pass the argument "--help" to the `sql_compare_file` module.

This module accepts a tab-delimited file as an input and calls the `sql_compare` module for each line in the file.

For Stored procedures, you may need to add `SET NOCOUNT ON; ` before your query to prevent strange errors from pyodbc.
You'll know you need this if Pyodbc throws some error like `preceding statement is not a query` or the cursor 
description comes back empty.  E.g. when saving sql results to excel you might get an error saying `NoneType is not
iterable`.  This occurs because the column headers are derived from the cursor description, and the cursor description 
will be blank if pyodbc doesn't think the statement is a query.  

### Templates
An excel template and example tab-delimited file are provided in the `example_file_input` directory.  

The excel template is provided for ease of use.  You can fill out the template then `Save As` a tab-delimited file 
and it should match the specifications required.  

The tab-delimited file is provided to show how an example of how to use the file input variation of the sql compare
module.

## Make executable
Compiling the modules to an executable is optional, but can help people who are not familiar with python 
use the command line tool.  It can also be useful for deploying to a server or shared drive.  

To deploy as an executable, I recommend using [pyinstaller](https://www.pyinstaller.org/).

### Compile Excel Compare Tool as Executable

Install pyinstaller, then compile the excel comparison module by executing the following command:

`pyinstaller compare.py -F -n compare_excel -i icon.ico`

1. The -F flag is for a one-file bundled executable
2. The -n flag gives the bundled app a name
3. The -i flag assigns an icon to the application

This will create a directory called "dist" in the project directory with containing your executable. Distribute the executable to your users.

On Windows, you may need to install some microsoft packages, like the VC++ redistributable package, prior to being able to compile the application.

### Compile SQL Compare Tool as executable

Install pyinstaller, then compile the excel comparison module by executing the following command:

`pyinstaller sql_compare.py -F -n sql_compare -i icon_sql.ico`

See above section for flags and meanings.

### Compile SQL File Input Compare Tool as executable

Install pyinstaller, then compile the excel comparison module by executing the following command:

`pyinstaller sql_compare_file.py -F -n file_sql_compare -i icon_sql_file.ico`

See above section for flags and meanings.

### Use via command line on Windows
The best way I've found to use this tool is to put the executable file in a folder that is part of your Windows path. 
By doing so, you can call the program easily by just typing the name.

To do this, I created a folder in my local drive called "CustomBatch" for all my commands.  Then I add the executable, 
or any other batch file, to the folder.  Finally I add the folder to my Windows path.  

You do this by opening your Advanced System Settings, find "Path" in your System variables, click Edit, 
then create a new entry to the list for your newly created folder 

![Pciture of environment variables](https://i.imgur.com/ESJabRO.png)
 
### Troubleshooting

If you get an error like this then you need to install the Microsoft Visual C++ package.

![System Error
Q The program can’t start because api-ms-win-crt-stdio-l1-1-OdIl is
‘ missing from your computer. Try reinstalling the program to fix this
problem.](https://i.imgur.com/eTgqVN4.png)


See [this link](http://www.thewindowsclub.com/api-ms-win-crt-runtime-l1-1-0-dll-is-missing) for more information.



#### Download links for Visual C++ package

Depending on your machine, you may need one or the other

[32 bit download](http://www.microsoft.com/en-gb/download/details.aspx?id=5555)

[64 bit download](http://www.microsoft.com/en-us/download/details.aspx?id=14632)
