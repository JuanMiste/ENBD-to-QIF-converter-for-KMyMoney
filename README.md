# ENBD to QIF converter for KMyMoney
 Script that converts Emirates NBD statements (2004 XLS files - xml based) to qif format ready for consumption by KMyMoney, including a simple recurring transaction classification function

Emirates NBD allows for downloading transactions history in a xls file (xml based). I use this script to convert files to a qif file to manually import the transactions into KMyMoney, the awesome open source Personal Finance software. Try it out! 

## Automatic labelling of transactions
Recurring transactions are recognised and automatically assigned a category based on a string which can be anywhere in the transaction. The list of strings and the accompanying Payee and Category for KMyMoney are listed in an .csv file for easy editing. The format is:
The string from statement,payee,category

## Run standalone
The easiest way to use this script is to edit enbd2qif.py itself and run it standalone:
- copy the statement and bankstatementmapping.csv to the script folder 
- Edit the enbd2qif.py file to fill in the enbd IBAN number/credit card name and the KMyMoney account name.
- you can add text to list (text_to_remove), these will be removed from descrption before writing unkown payee to text file
- Run the enbd2qif.py file.
- the script will generate text file with unmapped payees, edit the bankstatementmapping.csv file to automatically categorise transactions or set variable to False to skip the mapping.
- rerun the script
- Import the resulting qif file in KMyMoney.

## License
This script is provided under the MIT License. Please see the license file for more details.
