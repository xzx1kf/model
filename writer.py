import xlsxwriter
import xml.etree.ElementTree as ET
from sheet import Sheet

hardware_dict = {}

def create_hardware_profile_sheet(wrkbk):
    hardware_sheet = wrkbk.add_worksheet('Input Hardware Profile')

    headers = ['Hardware', 'Host Name', 'No. of cpu (desired)', 'CPU rating (GHz)', 'memory (GB) (Desired)', 'disk capacity (GB)']
    columns = ['hostname', 'cpu', 'cpu_rating', 'memory', 'disk']

    write_headers(hardware_sheet, headers, 5)
    return write_data(hardware_sheet, 7, 0, 'hardware', columns)


def create_input_resource_details_sheet(wrkbk, hardware_dict):
    sheet = wrkbk.add_worksheet('Input Resource Details')

    headers = ['Resource Code', 'Hardware', 'Description', 'Unit']
    columns = ['hardware', 'description', 'unit']

    write_headers(sheet, headers, 5)
    write_data(sheet, 7, 0, 'resource', columns)


def write_headers(sheet, headers, row):

    # Define the bold format.
    bold = workbook.add_format({'bold': True})

    # Write the header in bold.
    col = 0
    for header in headers:
        sheet.write(row, col, header, bold)
        col += 1


def write_data(sheet, row, col, data_root, columns):

    # Process xml data and write it to the sheet.
    tree = ET.parse('data.xml')
    root = tree.getroot()

    for data in root.findall(data_root):
        sheet.write(row, col, data.get('name'))

        if data_root == 'resource':
            name = data.find('hardware').text
            row_col = hardware_dict[name]
            rowH, colH  = str.split(row_col, ':')

            sheet.write(row, col, "='Input Hardware Profile'!A8")

        # Do I need to keep track of the above row and col index and write it out somewhere
        # for the input resource details sheet.
        # if 'data_root' == 'hardware' then write out a dictionary?

        if data_root == 'hardware':
            hardware_dict[data.get('name')] = str(row) + ':' + str(col)

        for column in columns:
            sheet.write(row, columns.index(column)+1, data.find(column).text)
        row += 1

    return hardware_dict

if __name__ == '__main__':
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('model.xlsx')
    business_transactions_sheet = workbook.add_worksheet('Input Business Transactions')
    it_transactions_sheet = workbook.add_worksheet('Input IT Transactions')

    # Create styles
    bold = workbook.add_format({'bold': True})

    header_row = 7

    btHeaders = ['Business Transaction Name',
                'Transaction Description',
                'Business Volumes',
                'Frequency',
                'Notes']
    btSheet = Sheet(business_transactions_sheet)
    btSheet.add_header(btHeaders)
    btSheet.write_headers(header_row, bold)

    itHeaders = ['Business Transaction',
                'IT Transaction Name',
                'IT Transaction Description',
                'Qty per transaction',
                'Transaction Rating',
                'TPS',
                'Notes']
    itSheet = Sheet(it_transactions_sheet)
    itSheet.add_header(itHeaders)
    itSheet.write_headers(header_row, bold)

    # Parse the xml data.
    tree = ET.parse('data.xml')
    root = tree.getroot()

    itrow = 0
    for row, bt in enumerate(root.findall('business_transaction')):
        rowdata = []
        rowdata.append(bt.get('name'))
        rowdata.append(bt.find('description').text)
        rowdata.append(bt.find('volume').text)
        rowdata.append(bt.find('frequency').text)
        rowdata.append(bt.find('notes').text)

        btSheet.track_column(0, 0)
        mappingName = rowdata[0]

        btSheet.track_column(0, 2)
        btSheet.write_row(row + header_row, rowdata)

        for it in bt.findall('it_transaction'):

            print itrow
            rowdata = []
            btSheet.print_mappings()
            busMappingName = mappingName + '#0'
            data = itSheet.write_mapping(busMappingName, btSheet)
            rowdata.append(data)
            rowdata.append(it.get('name'))
            rowdata.append(it.find('description').text)


            qty = 0
            try:
                qty = it.find('qty').text
                rowdata.append(float(qty))
            except:
                rowdata.append('')

            rowdata.append(1)
            busMappingName = mappingName + '#2'
            volume = itSheet.write_mapping(busMappingName, btSheet)
            print volume+'*D'+str(itrow+header_row)+'E'+str(itrow+header_row)+'/60/60'
            volume = volume+'*D'+str(itrow+header_row+1)+'*E'+str(itrow+header_row+1)+'/60/60'

            rowdata.append(volume)
            itSheet.write_row(itrow + header_row, rowdata)
            itrow += 1

            # Start of the IT to Resource Mapping Sheet
            for i in it.findall('resources'):
                for n in i.findall('name'):
                    print n.text
    workbook.close()
