import xlrd
from datetime import datetime


class ExtractData:
    """Extract data from an excel sheet"""

    def __init__(self, filename):
        """Initialize the required filename and call utility functions to open excel document"""

        self.excel_filename = filename
        self.workbook, self.sheet = self.make_workbook()
        self.nrows, self.ncols = self.sheet.nrows, self.sheet.ncols

    def make_workbook(self):
        """Utility function to open the excel sheet"""

        fp = xlrd.open_workbook(self.excel_filename)
        return fp, fp.sheet_by_index(0)

    def convert_date_to_str(self, date, date_format):
        """Convert excel date into the given string format"""

        excel_datetime = datetime(
            *xlrd.xldate_as_tuple(
                date,
                self.workbook.datemode,
            )
        ).strftime(date_format)

        return excel_datetime




    def process_header(self):

        """process the header part of the file and store data in a dict instance variable"""

        self.data = {}
        required_fields = [
            "Quote Number",
            "Date",
            "Ship To",
            "Ship From",
            "Name",
        ]

        for row in range(self.nrows):
            for col in range(self.ncols):
                cell_value = self.sheet.cell(row, col).value

                if cell_value in required_fields or "Name" in str(cell_value):

                    next_cell_value = self.sheet.cell(row, col + 1).value

                    if cell_value == "Date":
                        next_cell_value = self.convert_date_to_str(
                            next_cell_value, "%Y-%m-%d"
                        )

                    if "Name" in cell_value:

                        next_cell_value = cell_value.split(":")[1]
                        cell_value = "Name"

                    self.data[cell_value] = next_cell_value

        return self.data

    

    def process_columns(self):

        '''Extracts data from columns and stores it in a list instance variable'''

        self.column_data = []
        required_fields = [
            "LineNumber",
            "PartNumber",
            "Price",
            "Description",
        ]

        #Find the row number where column data starts
        for r in range(self.nrows):
            for c in range(self.ncols):
                cell_value = self.sheet.cell(r, c).value
                if cell_value in required_fields:
                    column_info_row = r
                    break
            else:
                continue
            break

        
        #start extracting values from the data
        for row in range(column_info_row, self.nrows):
            item = {}

            for col in range(0, self.ncols):

                cell_value = self.sheet.cell(row, col).value
                column_name = self.sheet.cell(column_info_row, col).value

                if row > column_info_row:

                    if (
                        not (cell_value == None or cell_value == "")
                        and column_name in required_fields
                    ):
                        item[column_name] = cell_value

            if len(item) > 0:
                self.column_data.append(item)

        return self.column_data

    def process_entire_sheet(self):
        '''process both header and column and return the result'''

        self.process_header()
        self.process_columns()
        self.data["Items"] = self.column_data

        return self.data






class TestExtractData:
    """Runner Class"""

    def __init__(self, filename, expected_output):
        """initialize """

        self.filename = filename
        self.expected_output = expected_output


    def compare_outputs(self ):
        '''Compare actual output with expected'''
        if len(self.actual_output) == len(self.expected_output) and all(
            x in self.expected_output for x in self.actual_output
        ):
            return True
        else:
            False

    def test(self):
        """Run the test"""

        ps = ExtractData(self.filename)
        self.actual_output = ps.process_entire_sheet()
        return self.compare_outputs()

        
            


expected_output = {
    "Quote Number": 98765.0,
    "Date": "2019-01-01",
    "Ship To": "USA",
    "Name": "Rapahel Epstein",
    "Items": [
        {
            "LineNumber": 1.0,
            "PartNumber": "ABC",
            "Description": "Very Good",
            "Price": 200.2,
        },
        {
            "LineNumber": 2.0,
            "PartNumber": "DEF",
            "Description": "Not so good",
            "Price": 100.1,
        },
    ],
}


filename = "./docs/Python Skill Test.xlsx"


ted = TestExtractData(filename, expected_output)

print(ted.test()) #Prints True or False depending upon if test case has passed or failed
