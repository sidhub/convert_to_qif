# Conversion to Quicken Interchange Format (QIF) from Excel
With this utility it is possible to convert the transaction in excel to QIF format which can be then imported into Quicken.

# Configuration file
To use the utility the configuration file needs to be configured. The utility takes the configuration as input behavior. The configuration file recommended to be named with extension
.properties
It should always contain two section.
- Section [general] :
    This section should contain the configuration to map the excel columns and the corresponding
    data that it contains. The following keys are mandatory configuration
    1) type : Account type specific in Quicken e.g. CCard, you can export your existing account from Quicken into QIF file to see what value it exports. The same can be then used.
    2) col_no_t_date : Column number which contains the transaction date (Note: Column index start with 1 as first column)
    3) col_no_desc : Column number which contains Description of the transaction. The keywords in this column will be used to match the provided keywords to identify the Quicken Category and corresponding memo text
    4) col_no_t_type : Column number which can identify whether the transaction is debit or credit, this will be always in context with the property 'col_t_type_credit_text'
    5) col_t_type_credit_text : The text that should be in the column (identified by col_no_t_type) that can be used to identify whether the transaction is a credit transaction, anything else will be considered debit
    6) col_t_type_amt : Column number in which the transaction amount exists
    7) default_category : Existing Quicken category if no match found it will be set e.g. Misc

- Section [mapping] :
    This section is to be configured by user based on the regular transaction text he uses for identifying the category. The keys under this mapping are the keywords from the transaction text which can be more than one comma separated whereas the values are the Categories in Quicken which you have created. It is also possible to provide a memo text corresponding to the category by adding the values in square brackes [] that follows the category
    e.g.
    Dominos,Subway,McDonald=Dining:Restaurant
    Walmart,Kmart,Macy=Household
    pharmacy,medicine= Medical:Pharmacy [some text that you want to add in memo field]

# Program execution
Execute the file toqif.py in Python environment to start.
The program requires three input from user
1) Property file name
2) Input transaction excel name
3) Output file name

By default the program will scan .properties file for Property and .xlsx for Input transaction in the current location where the program file exists and show them to user to select one. The user may select the index number shown corresponding to the file to select it or any other number not in index to allow to enter your own. i.e. if the entered index does not match the existing then the program will ask to enter the path.
For Output file name, it will ask to enter the path first if you just enter without entering path then the program will write the file named output.QIF file in the same folder.

# Importing QIF file
One can use built in the Import QIF file to import the new transactions from the QIF file.

_Note : The Quicken may not allow to import from the external QIF file into the accounts of type bank, credit card etc. In such you can create a new account of CASH type. Import them into this account, after that you may review the import and changes the category as desired if the mapping didnt resulted properly and after review select all the transaction and right click to select Move Transaction to the desired Bank, Credit card account._


# Sample files
Two sample files included:
1) sample.properties : property file sample
2) sample_input.xlsx : Sample input transaction excel downloaded from bank or card statement and modified to remove merge columns and other formatting

# Possible Error
It may be possible that the program fails in reading the transaction date column due to date format. In this case one may change the date format in the program itself. Please change the variable name *XLS_DATE_FORMAT* in toqif.py file