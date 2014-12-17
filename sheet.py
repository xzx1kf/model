import string
from string import maketrans

class Sheet():
    def __init__(self, sheet):
        self.sheet = sheet
        self.headers = []
        self.mappings = {}

    def add_header(self, header):
        if isinstance(header, basestring):
            self.headers.append(header)
        else:
            self.headers = header

    def write_headers(self, row):
        for header in self.headers:
            self.sheet.write(row, self.headers.index(header), header)

    def write_row(self, row, data, track_column = None):
        print data
        col = 0
        for cell in data:
            self.sheet.write(row, col, cell)

            print cell, row, data.index(cell)
            if track_column is not None:
                for i in track_column:
                    if data.index(cell) == i:
                        self.create_mapping(str(data[0])+'#'+str(i), row, i)
            col += 1

    def create_mapping(self, name, row, col):
        # Add one to the row because excel is not zero indexed.
        self.mappings[name] = str(row+1) + ':' + str(col)

    def get_mapping(self, name):
        return self.mappings[name]

    def get_name(self):
        return self.sheet.get_name()

    def write_mapping(self, name, other_sheet):
        row, col = other_sheet.get_mapping(name).split(':')
        col = chr(int(col) + 65)                                                    # Map column 0-25 onto A-Z
        return "='%s'!%s%s" % (other_sheet.get_name(), col,  row)

    def print_mappings(self):
        for mapping in self.mappings:
            print "%s, %s" % (mapping, self.mappings[mapping])


if __name__ == '__main__':
    import xlsxwriter
    import xml.etree.ElementTree as ET
    workbook = xlsxwriter.Workbook('model.xlsx')
    business_transactions_sheet = workbook.add_worksheet(\
            'Input Business Transactions')
    sheet = Sheet(business_transactions_sheet)
    headers = ['Business Transaction', 'Transaction Description', 'Business Volumes', 'Frequency', 'Notes']
    #sheet.add_header('Business Transaction')
    #sheet.add_header('Transaction Description')
    #sheet.add_header('Business Volumes')
    #sheet.add_header('Frequency')
    #sheet.add_header('Notes')
    sheet.add_header(headers)

    sheet.write_headers(6)
    sheet.write_row(7, ['Process Enquiry', 'Process Enquiry / Application', 80, 'per hour', ''], [0, 2])

    sheet.print_mappings()
    it_trans_sheet = workbook.add_worksheet(\
            'Input IT Transactions')
    itsheet = Sheet(it_trans_sheet)

    itsheet.add_header('Business Transaction')
    itsheet.add_header('IT Transaction Name')
    itsheet.add_header('IT Transaction Description')
    itsheet.add_header('Qty')
    itsheet.add_header('Transaction Rating')
    itsheet.add_header('TPS')

    itsheet.write_headers(6)

    #row, col = sheet.get_mapping('Process Enquiry').split(':')
    #itsheet.write_row(7, ["='Input Business Transactions'!A8", 'IT Transaction', 10], 1)
    print itsheet.write_mapping('Process Enquiry#0', sheet)
    itsheet.write_row(7, [itsheet.write_mapping('Process Enquiry#0', sheet), 'Process Enquiry / Application', 'Process Enquiry / Application', 1, 1, itsheet.write_mapping('Process Enquiry#2', sheet)])

    workbook.close()

