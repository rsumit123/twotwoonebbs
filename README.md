# Assignment provided by TwoTwo1BBS

## Extract Excel Data

### Problem Statement

- The objective is to read, and parse a given Excel file, using the “xlrd” module. Note:- If you are
  using Python3 then use xlrd==1.2.0
- The result should be nicely printed as a Dict (JSON)
- The input Excel has two parts: The header which is built like a form with labels and values,
  followed by rows with several columns.
- Each label should be a key in the output Dict, with the corresponding value.
- The rows should be a list element in the DIct, where each item should have the column header
  as the key.
- There is a list of fields that should be extracted from the header part, and a list of fields from
  the columns. Fields not on the list should be ignored.

- Date values (label has the string somewhere in it) should be converted to the forma “yyyy- mm-dd”
- Header list to be extracted:
- The code should look for form field labels called “Quote Number”, “Date”, “Ship To”, “Ship From”, “Name”.
- The value is usually to the right of the label.
- There is one exception, and that is “Name”, that both label and value are stored in the same Excel cell with a colon as a separator (like “Name: StrataVAR”).
- Item columns to be extracted: LineNumber, PartNumber, Description, Price. Other columns should be ignored.
- Incase an expected field is not found, print a detailed error message to the console, continue the process nevertheless.

### Solution

- Created one class for problem and one class for running and testing
- Divided the problem into two parts for extracting header info and extracting column info
- Followed the above steps to ensure all the cases are covered
- Tested to ensure that expected output == actual output

## Compare and process two lists (Bom and Disti)

### Problem Statement

- Create a 3rd list, with 5 fields: bom_pn, bom_qty, Dsti_pn, Dsti_qty, Error Flag.
- All items from the BoM will appear as is and based on the same order of the input BoM.
- Against each bom_pn there should be a Dsti_pn (if one exists). If Disti PartNumber is missing a Part Number in Disti, the Disti side should be left blank and Error Flag should be checked
- It is possible that Disti will be missing PN or BOM will be missing some Part Numbers
- If Disti Quantity is bigger than BoM Quantity (for any given Part Number), Disti line should be split to match the Quantity of the BoM, and the remainder should be aligned with another BoM line, if exists.
- If there is another record with same part number, the above should be repeated.
- The joint list should be as similar as possible in term of quantities.
- You may end up with some lines that do not have a Disti values, and vice-a-versa.
- No matter what, the overall quantities per Part Number on BoM side in the Unified List output table should be the same as the original BoM. Same for Disti.
- The Disti Side, as a whole, should include ALL items from original Disti, even if lines were split.

### Solution

- Created one class for problem and one class for running and testing
- Copied all the values from BOM list to unified list
- Iterated through the bom list searching of disti quantity w.r.t to BOM Part number
- Handled three cases namely
  - Bom quantity is greater than disti
  - disti quantity is greater than bom
  - both quantities are equal
- Handled the cases where remainder disti items are left (added them with error flag x)
- Tried to handle all the cases given in the above problem statement
- Tested to ensure that expected output == actual output

## Dependencies

- Python3.9
- xlrd==1.2.0

## Installation

- clone the project to your machine
- cd into the project directory (ex: `cd twotwoonebbs`)
- run pipenv shell (if pipenv is not installed run `pip install pipenv` first)
- run pipenv install to install all the dependencies
- start executing python files extract_data and bomdisti by `python extract_data.py` and `python bomdisti.py`

Please contact me at rsumit123@gmail.com for more details
