import sys
from bs4 import BeautifulSoup
from pathlib import Path
import openpyxl
import datetime
import csv

statement_name = input('Please enter the statement name: ')

qiffile = Path(f'{statement_name.replace(".xls", "")}.qif') 

# Below list of strings, these are removed from the unmapped payees, add any other text to the list if needed 
text_to_remove = [" DUBAI ARE"," ABU DHABI ARE"]

# name of accounts/credit card in ENBD (for accounts in format "account_number(IBAN_number)")
enbd_aed_iban ="xxxxxxxxxxxxx(AExxxxxxxxxxxxxxxxxxxx1)"
enbd_usd_iban = "yyyyyyyyyyyyy(AExxxxxxxxxxxxxxxxxxxx2)"
enbd_ccard_name = "ENBD Credit Card Name"

# name of accounts/credit card in Kmymoney 
kmy_nbd_aed_name = "Current account - AED"
kmy_nbd_usd_name = "Current account - USD"
kmy_nbd_ccard_name = "U by Emaar"

MAP_FILE = Path('bankstatementmapping.csv')  # False

with open(statement_name, 'r') as f:
        xml_string = f.read()
statement_xml = BeautifulSoup(xml_string, 'xml')
transactions =[]

class StatementExtract:

    def check_type(statement_xml):
        statement_type = ""
        for row in statement_xml.find_all('Row'):
                    if (row['ss:AutoFitHeight']) == '0' and (row['ss:Height']) == '33.0':
                        for cell in row.find_all('Cell'):
                            trans_payee    = ""
                            trans_category = ""
                            if (cell['ss:Index']) == '5' and (cell['ss:StyleID'])== '28':
                                if cell.find('ss:Data'):
                                    if cell.find('ss:Data').text ==  enbd_aed_iban:
                                        statement_type = kmy_nbd_aed_name
                                    elif cell.find('ss:Data').text == enbd_usd_iban:
                                        statement_type = kmy_nbd_usd_name
                    elif (row['ss:AutoFitHeight']) == '0' and (row['ss:Height']) == '27.75':
                        for cell in row.find_all('Cell'):
                            trans_payee    = ""
                            trans_category = ""
                            if (cell['ss:Index']) == '1' and (cell['ss:StyleID'])== '28':
                                if cell.find('ss:Data'):
                                    if cell.find('ss:Data').text == enbd_ccard_name:
                                        statement_type = kmy_nbd_ccard_name
        return statement_type

    def load_xml(statement_xml):

        """
        convert xml data to a list of: Date, Payee, Category, Description and amount
        Expects the XML based Excel file (1997-2004) 
        """
        # if bank statement is for current account in AED or foreign currency
        if StatementExtract.check_type(statement_xml) == kmy_nbd_aed_name or StatementExtract.check_type(statement_xml) == kmy_nbd_usd_name:
            for row in statement_xml.find_all('Row'):
                if (row['ss:AutoFitHeight']) == '1':  
                    for cell in row.find_all('Cell'):
                        trans_payee    = ""
                        trans_category = ""
                        if (cell['ss:Index']) == '1':
                            if cell.find('ss:Data'):
                                trans_date = cell.find('ss:Data').text
                                if len(trans_date) == 11:
                                    date_time_obj = datetime.datetime.strptime(trans_date, '%d %b %Y')
                                    trans_date =  date_time_obj.strftime('%d.%m.%Y')
                        elif (cell['ss:Index']) == '3': 
                            if cell.find('ss:Data'):
                                trans_short_desc = cell.find('ss:Data').text
                        elif (cell['ss:Index']) == '11': 
                            if cell.find('ss:Data'):
                                trans_amount = cell.find('ss:Data').text
                        elif (cell['ss:Index']) == '13': 
                            if cell.find('ss:Data'):
                                trans_credit = cell.find('ss:Data').text
                                if trans_amount == "":
                                    trans_amount = trans_credit
                        elif (cell['ss:Index']) == '16': 
                            if cell.find('ss:Data'):
                                trans_balance = cell.find('ss:Data').text

                            trans_desc = " Date= "+str(trans_date)+" Desc= "+str(trans_short_desc)+" Amount= "+str(trans_amount)+" Balance= "+str(trans_balance)
                        
                            transactions.append([trans_date,trans_payee,trans_category,trans_desc,trans_amount])
            return transactions
        # if bank statement is for credit card
        elif StatementExtract.check_type(statement_xml) == kmy_nbd_ccard_name:

            for row in statement_xml.find_all('Row'):
                if (row['ss:AutoFitHeight']) == '0' and (row['ss:Height']) == '22.5':  
                    for cell in row.find_all('Cell'):
                        trans_payee    = ""
                        trans_category = ""
                        if (cell['ss:Index']) == '1':
                            if cell.find('ss:Data'):
                                trans_date = cell.find('ss:Data').text
                                if len(trans_date) == 11:
                                    date_time_obj = datetime.datetime.strptime(trans_date, '%d %b %Y')
                                    trans_date =  date_time_obj.strftime('%d.%m.%Y')
                        elif (cell['ss:Index']) == '10': 
                            if cell.find('ss:Data'):
                                trans_short_desc = cell.find('ss:Data').text
                        elif (cell['ss:Index']) == '16': 
                            if cell.find('ss:Data'):
                                trans_card = cell.find('ss:Data').text
                        elif (cell['ss:Index']) == '19': 
                            if cell.find('ss:Data'):
                                trans_amount = cell.find('ss:Data').text
                            trans_amount =  trans_amount.replace("AED ", "") 

                            trans_desc = " Date= "+str(trans_date)+" Desc= "+str(trans_short_desc)+" Amount= "+str(trans_amount)+" Card Type= "+str(trans_card)
                        
                            transactions.append([trans_date,trans_payee,trans_category,trans_desc,trans_amount])
            
        return transactions

KMYMONEYACCOUNT =  StatementExtract.check_type(statement_xml)

def load_mapdict(filename_mapfile=MAP_FILE):
    mapping = {}
    # Open the CSV file
    with open(filename_mapfile, 'r') as mapfile:
        print(f'Loading mapping from  {filename_mapfile}')
        reader = csv.reader(mapfile)
    
        for row in reader:
            if (row[0] is not None) and row[0] != ' ':
                mapping[row[0]] =  (row[1], row[2])
    
    mapfile.close()
 
    return mapping

def map_transactions(transactions, mapping):
    """
    Compare all transactions to a dictionary with payees and categories
    """
    mappedtransactions = []
    mapcounter = 0
    payee = "Payee to be checked"
    category = "Category to be checked"
    amount = ""
    payee_to_check =[]

    for i in transactions:
        if len(i) > 0:
            if i[0] is not None:
                date = i[0]
            if i[1] is not None:    
                payee = i[1]
            if i[2] is not None:    
                category = i[2]    
            if i[3] is not None:      
                desc = i[3]
            if i[4] is not None:    
                amount = i[4]
      
            for identifier in mapping.keys():
                if identifier.lower() in desc.lower():
                    payee, category = mapping[identifier]
                    mapcounter += 1
                    break 
                else:
                    payee = "Payee to be checked"
                    category = "Category to be checked"
                    
            mappedtransactions.append([date, payee, category, desc, amount])

   # print(f'Mapped {mapcounter} out of {len(transactions)} ' +
   #      f'transactions ({100.*mapcounter/len(transactions):.0f}%)')

    return mappedtransactions

def write_transactions_to_qif(transactions, qiffile, verbose=False):
    """
    Take a list of transaction dictionaries and write to QIF file
    transactions: list of transactions
    qiffile: filename to write contents to
    """
    with open(qiffile, 'w') as f:
        if KMYMONEYACCOUNT == kmy_nbd_ccard_name:
        # Write opening statement for KMyMoney if ccard
            f.write('!Type:CCard\n')
            f.write('CX\n')
            f.write('POpening Balance\n')
            f.write(f'L[{KMYMONEYACCOUNT}]\n') 
            f.write('^\n')
        else:
        # Write opening statement for KMyMoney if account
            f.write('!Account\n')
            f.write(f'N{KMYMONEYACCOUNT}\n') 
            f.write('TBank\n')
            f.write('^\n')
            f.write('!Type:Bank\n')
            f.write('^\n')

        # Loop over transactions and write to file
        for date, payee, category, desc, amount in transactions:
            f.write(f'D{date}\n')
            f.write(f'T{amount}\n')
            if payee != '':
                f.write(f'P{payee}\n')
            if category != '':
                f.write(f'L{category}\n')
            f.write(f'M{desc}\n')
            f.write('^\n')

def write_unkown_payees_to_txt(mapped_transactions):
    """
    If the payee is unmapped, the name of payee is written to TXT file ( after extracting it from the desc and removing dates and other texts)
    add this to the mapping file and re-run the script, at the end the TXT file should be empty.
    """    
    trans_payee_found = 0
    trans_payee_not_found = 0
    payee_not_found = 0
    payee_to_check =[]
    for i in mapped_transactions:
        
        if i[1]=="Payee to be checked":
            payee = i[3].split(" Amount= ",1)[0]
            payee = payee[24:]
            for text in text_to_remove:
                payee = payee.replace(text, "")
            payee_to_check.append(payee)
            trans_payee_not_found = trans_payee_not_found +1
        else:
            trans_payee_found = trans_payee_found +1
    file = open(f'{statement_name.replace(".xls", "")}.txt','w')

    for i in set(payee_to_check):
        file.write(i+"\n")
        payee_not_found = payee_not_found + 1
    file.close()
    print(f'Transactions with payee found =  {trans_payee_found}')
    print(f'Transactions with unkown Payee = {trans_payee_not_found}')
    print(f'Unkown Payee = {payee_not_found}')
    print(f'File with Unkown Payee = {file.name}')
    


def run():
    dic_items= load_mapdict()
    trans_list = StatementExtract.load_xml(statement_xml)
    mapped_transactions = map_transactions(trans_list, dic_items) 
    write_transactions_to_qif(mapped_transactions,qiffile)
    write_unkown_payees_to_txt(mapped_transactions)
    print(f"QIF account is: *{KMYMONEYACCOUNT}*")

if __name__ == "__main__":
    run()