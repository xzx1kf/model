class Sheet():
    def __init__(self, sheet):
        self.sheet = sheet
        self.headers = []
        self.mappings = {}
        self.track_columns = []

    def add_header(self, header):
        if isinstance(header, basestring):
            self.headers.append(header)
        else:
            self.headers = header

    def write_headers(self, row, format = None):
        # Becasue excel sheets are not 0 indexed need to subtract 1 from the row.
        row -= 1
        for header in self.headers:
            if format is None:
                self.sheet.write(row, self.headers.index(header), header)
            else:
                self.sheet.write(row, self.headers.index(header), header, format)

    def write_row(self, row, data):
        # enumerate simply allows access to both the index and the value in data list.
        for column, value in enumerate(data):
            self.sheet.write(row, column, value)
            # track_columns is a list of tuples, the if statement checks the second
            # item in the tuple (as that is the column index to be tracked) to see
            # if it should be tracked.
            if column in [x[1] for x in self.track_columns]:
                self.create_mapping(data[x[0]]+'#'+str(column), row, column)

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

    def track_column(self, key, column):
        """Track this column."""
        #column = chr(int(column) + 65)
        self.track_columns.append((key, column))


if __name__ == '__main__':
    import xlsxwriter
    #import xml.etree.ElementTree as ET
    workbook = xlsxwriter.Workbook('model.xlsx')
    business_transactions_sheet = workbook.add_worksheet(\
            'Input Business Transactions')
    sheet = Sheet(business_transactions_sheet)
    headers = ['Business Transaction', 'Transaction Description', 'Business Volumes', 'Frequency', 'Notes']
    sheet.add_header(headers)

    sheet.write_headers(6)
    sheet.track_column(0, 0)
    sheet.track_column(0, 2)
    sheet.write_row(7, ['Process Enquiry', 'Process Enquiry / Application', 80, 'per hour'])


    sheet.print_mappings()
    #""""
    it_trans_sheet = workbook.add_worksheet( 'Input IT Transactions')
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
    #"""
    workbook.close()

