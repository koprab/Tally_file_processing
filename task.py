import pandas as pd
import xlsxwriter
from xlsxwriter.exceptions import FileCreateError
import xml.etree.ElementTree as ET
from datetime import datetime
import logging
import os

logging.basicConfig(level=logging.INFO, format='%(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('process_file')
_file_name = './invalid_2.xml'
old_date_format = '%Y%m%d'
new_date_format = '%d-%m-%Y'

column_header = ['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount',
               'Amount', 'Particulars', 'Vch Type', 'Amount Verified']
data_list = []


def save_file_to_xls(data, col, file_name):
    """
    Method for saving file to xlsx
    convert it to DataFrame
    return : bool
  """
    saved = False
    try:
        print(f'[INFO] Trying to save file using Pandas ')
        if len(data) > 0 and len(col) > 0 and file_name is not None:
            df = pd.DataFrame(data, columns=col)  # dataframe can be used for further modifications
            if len(df) > 0:
                writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sample', index=False)
                writer.save()
                print(f'[INFO] File saved as : {os.path.basename(file_name)} with : {len(data)} lines ')
                saved = True
    except xlsxwriter.exceptions.FileCreateError as ex:
        logger.error(f'Error in saving file : ', ex)
        saved = False
    return saved


def save_to_file_using_xlsxwritre(data, col, file_name):
    """
    :param data: List of Lines
    :param col: columns header
    :param file_name: Name of file
    :return: bool
    """
    saved = False
    try:
        print(f'[INFO] Trying to save file using xlsxwriter')
        if len(data) > 0 and len(col) > 0 and file_name is not None:
            if os.path.exists(file_name):
                os.remove(file_name)
            with xlsxwriter.Workbook(file_name) as workbook:
                sheet = workbook.add_worksheet('Sample')
                header_format = workbook.add_format({'bold': True})
                for head_num, column_name in enumerate(col):
                    sheet.write(0, head_num, column_name,header_format)
                for index, row in enumerate(data):
                    sheet.write_row(index+1, 0, row)
                    saved = True
                print(f'[INFO] File saved as : {os.path.basename(file_name)} with : {len(data)} lines ')
    except FileCreateError as ex:
        logger.error(f'Error in saving file :', ex)
        saved = False
    return saved


def get_ref_amount_sum(child_element):
    """
    :param child_element: List of Elements
    :return: sum of amount
    """
    sum = 0.0
    for ch in child_element:
        value = ch.find('ISDEEMEDPOSITIVE').text
        if value == 'No':
            bills = ch.findall('.//BILLALLOCATIONS.LIST')
            for b in bills:
                sum += float(b.find('AMOUNT').text)
    return round(float(sum), 2)


def process_file():
    """
    :return: None
    """
    try:
        tree = ET.parse(_file_name)
        root = tree.getroot()
        print(f'[INFO] Processing File : {os.path.basename(_file_name)}')
        vouchers = root.findall('.//VOUCHER[@VCHTYPE="Receipt"]')
        if len(vouchers) > 0:
            for entry in vouchers:
                unformatted_date = datetime.strptime(entry.find('DATE').text, old_date_format)
                formatted_date = unformatted_date.strftime(new_date_format)
                parent_transaction_type = 'Parent'
                vch_no = entry.find('VOUCHERNUMBER').text
                parent_ref_no = 'NA'
                parent_ref_type = 'NA'
                parent_ref_date = 'NA'
                debtor = ' '.join([x.capitalize() for x in entry.find('PARTYLEDGERNAME').text.split(' ')])
                parent_ref_amount = 'NA'
                parent_particulars = debtor
                vch_type = 'Receipt'
                amount_verified = ''
                child_elements = entry.findall('ALLLEDGERENTRIES.LIST')
                if len(child_elements) >= 1:
                    total_amount = get_ref_amount_sum(child_elements)
                    # amount conversion to float
                    # parent_amount = float(entry.find('./ALLLEDGERENTRIES.LIST/AMOUNT').text)  # Amount header
                    parent_amount = total_amount
                    # logger.info(f'Parent_amount {parent_amount} & total_amount : {total_amount}')
                    if total_amount > 0:
                        if parent_amount == total_amount:
                            amount_verified = 'Yes'
                        else:
                            amount_verified = 'No'

                    data_list.append(
                        [formatted_date, parent_transaction_type, vch_no, parent_ref_no, parent_ref_type, parent_ref_date,
                         debtor, parent_ref_amount, f'{parent_amount:.2f}', parent_particulars, vch_type, amount_verified])

                    for child in child_elements:

                        other_or_child = child.find('ISDEEMEDPOSITIVE').text
                        if other_or_child == 'No':  # Child entry
                            child_debtor = ' '.join([x.capitalize() for x in child.find('LEDGERNAME').text.split(' ')])
                            child_particulars = child_debtor
                            child_recipt_type = 'Receipt'
                            child_amount_verified = 'NA'
                            bill_lists = child.findall('.//BILLALLOCATIONS.LIST')
                            for bill in bill_lists:
                                child_ref_no = bill.find('NAME').text
                                child_ref_type = bill.find('BILLTYPE').text
                                child_ref_amount = float(bill.find('AMOUNT').text)
                                child_amount = 'NA'
                                child_trans_type = 'Child'
                                child_ref_date = ''
                                data_list.append(
                                    [formatted_date, child_trans_type, vch_no, child_ref_no, child_ref_type, child_ref_date,
                                     child_debtor, f'{child_ref_amount:.2f}', child_amount, child_particulars, child_recipt_type,
                                     child_amount_verified])
                        elif other_or_child == 'Yes':  # other entry
                            other_debtor = child.find('LEDGERNAME').text
                            other_particular = other_debtor
                            other_amount = float(child.find('AMOUNT').text)
                            other_trans_type = 'Other'
                            other_ref_amount = 'NA'
                            other_ref_no = 'NA'
                            other_ref_type = 'NA'
                            other_amt_verified = 'NA'
                            other_ref_date = 'NA'
                            data_list.append(
                                [formatted_date, other_trans_type, vch_no, other_ref_no, other_ref_type, other_ref_date,
                                 other_debtor, other_ref_amount, f'{other_amount:.2f}', other_particular, vch_type, other_amt_verified])

                else:
                    print(f'[INFO] No Child Entries are present in xml')
        else:
            print(f'[INFO] No voucher type Receipt is present in xml')
        if len(data_list) > 0:
            # save file using pandas
            # saved = save_file_to_xls(data_list, header_list, './Processed_file.xlsx')

            # save file using xlsxwriter
            saved = save_to_file_using_xlsxwritre(data_list, column_header, './Processed_file_1.xlsx')
            if not saved:
                logger.error(f'Error saving file')
            else:
                print(f'[INFO] File Processing Completed')
        else:
            print(f'[INFO] Nothing there to save')
    except Exception as e:
        print(e)
        logger.error('Exception occurred', e)


if __name__ == '__main__':
    process_file()
