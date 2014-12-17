import xlsxwriter
import xml.etree.ElementTree as ET
import sys

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


# Create an new Excel fiel and add a worksheet.
workbook = xlsxwriter.Workbook('model.xlsx')
business_transactions_sheet = workbook.add_worksheet('Input Business Transactions')

# Create the Input IT Transactions Sheet
it_transactions_sheet = workbook.add_worksheet('Input IT Transactions')

# Set column width
business_transactions_sheet.set_column('A:A', 20)

# Create styles
bold = workbook.add_format({'bold': True})

# Write column headers
business_transactions_sheet.write('A6', 'Business Transaction Name', bold)
business_transactions_sheet.write('B6', 'Transaction Description', bold)
business_transactions_sheet.write('C6', 'Business Volumes', bold)
business_transactions_sheet.write('D6', 'Frequency', bold)
business_transactions_sheet.write('E6', 'Notes', bold)

# Write column headers
it_transactions_sheet.write('A6', 'Business Transaction', bold)
it_transactions_sheet.write('B6', 'IT Transaction Name', bold)
it_transactions_sheet.write('C6', 'IT Transaction Description', bold)
it_transactions_sheet.write('D6', 'Qty per transaction', bold)
it_transactions_sheet.write('E6', 'Transaction Rating', bold)
it_transactions_sheet.write('F6', 'TPS', bold)
it_transactions_sheet.write('G6', 'Notes', bold)

# Process xml data and write it to the sheet.
tree = ET.parse('data.xml')
root = tree.getroot()

row = 7
it_row = 7

for business_transaction in root.findall('business_transaction'):
    name = business_transaction.get('name')
    description = business_transaction.find('description').text
    volume = business_transaction.find('volume').text
    frequency = business_transaction.find('frequency').text
    notes = business_transaction.find('notes').text

    business_transactions_sheet.write('A' +  str(row), name)
    business_transactions_sheet.write('B' +  str(row), description)
    business_transactions_sheet.write('C' +  str(row), volume)
    business_transactions_sheet.write('D' +  str(row), frequency)
    business_transactions_sheet.write('E' +  str(row), notes)

    first_iteration = True

    for it_transaction in business_transaction.findall('it_transaction'):
        if first_iteration:
            it_transactions_sheet.write('A' + str(it_row), name)
        it_transactions_sheet.write('B' + str(it_row), it_transaction.get('name'))
        it_transactions_sheet.write('C' + str(it_row), it_transaction.find('description').text)
        try:
            qty = it_transaction.find('qty').text
            it_transactions_sheet.write('D' + str(it_row), qty)
            tps = float(qty) * float(volume) / 60 / 60
            it_transactions_sheet.write('F' + str(it_row), "='Input Business Transactions'!C"+str(row)+'*D'+str(it_row)+'/60/60')
            #it_transactions_sheet.write('F' + str(it_row), tps)
        except:
            pass

        it_row += 1
        first_iteration = False

    row += 1

hardware_dict = create_hardware_profile_sheet(workbook)
create_input_resource_details_sheet(workbook, hardware_dict)
workbook.close()
