This script downloads pictures (the list of search queries is taken from the xlsm file) from google search in png format.

Used third-party libraries such as selenium, requests, openpyxl, and several built-in libraries

Downloaded pictures are saved in the same folder where the script is run

To run the script, you must specify the following arguments

`required arguments:`

_**xlsm_file**_ path to the xlsm file with a list of search queries

`optional:`

_**-sheet_name**_ is the name of the sheet in the file containing the list of search queries (default Sheet1)

_**-img_num**_ how many images to download for each request (default 4)

_**-fc_club**_ flag indicating that searches are for the names of football clubs (default False)

In the xlsm file on the sheet, all search queries must be in column A, each query in a new cell (no separating characters, only the queries themselves)
