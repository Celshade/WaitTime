# OvenReader
[![GNU GPLv3 license](https://img.shields.io/badge/license-GPLv3-blue.svg)](https://github.com/Celshade/OvenReader/blob/master/LICENSE.txt)
[![built with Python3](https://img.shields.io/badge/built%20with-Python3-green.svg)](https://www.python.org/)

_Parse oven scheduling data and output details in a more useful format_

Welcome to WaitTime!

WaitTime is a program built to automate the process of parsing specific KPI
data from an Excel workbook. This Excel file contains weekly oven schedules -
each tab representing its own week, and each cell representing one hour.
WaitTime parses these spreadsheets and outputs the various KPIs in a simplified
format.

**\*\*Disclaimer\*\*** \
This project was built for parsing very specific data which is organized in
uniquely formatted Excel files. This format may not be universal, and the
program may need to be tuned accordingly for anyone using this code in their
own projects.

For the sake of privacy and compliance, all product information used in the
example file **_.\docs\example_schedule.xlsx_** (product names, product lot
numbers, and comments) has been removed.
***

## Requirements
1. _**Python [3.7+]**_: Can be downloaded ->[here](https://www.python.org/)<-
2. _**Pandas**_: Installable via `pip install pandas`.
3. _**Openpyxl**_: Installable via `pip install openpyxl`.

_Some older versions of python may work, but this has not been tested._
***

## Instructions
This program was built for the CLI. Simply download the package, and
run `py .\src\waittime.py` (from the package directory) to start the
program. The example file can be used for testing purposes, as follows:

(The **.\\** can be omitted if you're working from the package directory.) \
`py .\src\waittime.py` \
(At the prompt...) \
`.\docs\example_schedule.txt`

The program will prompt for a 'target week' - this refers to a tab in the
Excel file. Each week begins on Monday and can be entered in the format of
MM-DD-YY. For simplicity, I have given the weeks of 1-25-21 and 1-14-21 as
examples for testing.


**\*\*NOTE\*\*** \
_Users may need to run `python .\src\waittime.py` depending on their setup._
***

For the most recent developments, please see branch [DEV](https://github.com/Celshade/WaitTime/tree/dev).

Found a bug or care to drop some feedback? See the link below! \
_Comments & bugs -> https://www.github.com/Celshade/WaitTime/issues_

Peace

**//Cel**
