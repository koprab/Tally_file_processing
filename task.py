import pandas as pd
import xlsxwriter
from xlsxwriter.exceptions import FileCreateError
import xml.etree.ElementTree as ET
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('process_file')

tree = ET.parse('./Input.xml')
root = tree.getroot()

old_date_format = '%Y%m%d'
new_date_format = '%d-%m-%Y'

header_list = ['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount',
               'Amount',
               'Particulars', 'Vch Type', 'Amount Verified']
data_list = []
vouchers = root.findall('.//VOUCHER[@VCHTYPE="Receipt"]')


def save_file_to_xls(data, col, file_name):
    """
    Method for saving file to xlsx
    convert it to DataFrame
    return : None
  """

    try:
        if len(data) > 0:
            df = pd.DataFrame(data, columns=col)  # dataframe can be used for further modifications
            if len(df) > 0:
                writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sample', index=False)
                writer.save()
                print(f'File saved as {file_name} with : {len(data)} lines ')
    except xlsxwriter.exceptions.FileCreateError as ex:
        print(f'Error in saving file : ', ex)


def get_ref_amount_sum(child_element):
    """"
    Method to check for amount sum and set Amount Verified
    accordingly Either Yes| No
    return : sum of amount
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
    Process file

  """
    try:
        for entry in vouchers:
            unformatted_date = datetime.strptime(entry.find('DATE').text, old_date_format)
            formatted_date = unformatted_date.strftime(new_date_format)
            parent_transaction_type = 'Parent'
            parent_amount = float(entry.find('./ALLLEDGERENTRIES.LIST/AMOUNT').text)  # Amount header
            vch_no = int(entry.find('VOUCHERNUMBER').text)
            parent_ref_no = 'NA'
            parent_ref_type = 'NA'
            parent_ref_date = 'NA'
            debtor = ' '.join([x.capitalize() for x in entry.find('PARTYLEDGERNAME').text.split(' ')])
            parent_ref_amount = 'NA'
            parent_particulars = debtor
            vch_type = 'Receipt'
            amount_verified = ''
            child_elements = entry.findall('ALLLEDGERENTRIES.LIST')
            total_amount = get_ref_amount_sum(child_elements)

            # logger.info(f'Parent_amount {parent_amount} & total_amount : {total_amount}')
            if total_amount > 0:
                if parent_amount == total_amount:
                    amount_verified = 'Yes'
                else:
                    amount_verified = 'No'

            data_list.append(
                [formatted_date, parent_transaction_type, vch_no, parent_ref_no, parent_ref_type, parent_ref_date,
                 debtor,
                 parent_ref_amount, parent_amount, parent_particulars, vch_type, amount_verified])

            for child in child_elements:  # last entry is for other type

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
                             child_debtor
                                , child_ref_amount, child_amount, child_particulars, child_recipt_type,
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
                         other_debtor,
                         other_ref_amount, other_amount, other_particular, vch_type, other_amt_verified])

        save_file_to_xls(data_list, header_list, './Processed_file.xlsx')
    except Exception as e:
        print(e)


if __name__ == '__main__':
    process_file()
