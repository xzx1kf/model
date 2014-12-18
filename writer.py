import xlsxwriter
import xml.etree.ElementTree as ET
from sheet import Sheet

hardware_dict = {}

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
            volume = volume+'*D'+str(itrow+header_row+1)+'*E'+str(itrow+header_row+1)+'/60/60'

            rowdata.append(volume)
            itSheet.write_row(itrow + header_row, rowdata)
            itrow += 1

            # Start of the IT to Resource Mapping Sheet
            for i in it.findall('resources'):
                for n in i.findall('name'):
                    print n.text
    workbook.close()
