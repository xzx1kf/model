import xlsxwriter
import xml.etree.ElementTree as ET
import sys

headers = ['Business Transaction Name', \
            'Transaction Description', \
            'Business Volumes', \
            'Frequency', \
            'Notes']

def write_headers(sheet, headers, row):
    bold = workbook.add_format({'bold': True})  # Define the bold format.
    col = 0
    for header in headers:
        sheet.write(row, col, header, bold)     # Write the header in bold.
        col += 1

def write_data(sheet, data_root, row, col):
    tree = ET.parse('data.xml')
    root = tree.getroot()
    records = root.findall(data_root)

    for i in range(len(records)):
        sheet.write(row+i, col, records[i].get('name'))
        sheet.write(row+i, col+1, records[i].find('description').text)
        try:
            sheet.write_number(row+i, col+2, int(records[i].find('volume').text))
        except:
            pass
        sheet.write(row+i, col+3, records[i].find('frequency').text)
        sheet.write(row+i, col+4, records[i].find('notes').text)

if __name__ == '__main__':
    workbook = xlsxwriter.Workbook('model.xlsx')
    business_transactions_sheet = workbook.add_worksheet(\
            'Input Business Transactions')
    write_headers(business_transactions_sheet, headers, 7)
    write_data(business_transactions_sheet, 'business_transaction', 9, 0)
    workbook.close()
