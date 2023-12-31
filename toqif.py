'''''
Program to transform the xls into Quicken Interchange Format (QIF)
This output can then be used to import directly into Quicken
 
Normally those banks/credit card statements that cannot be directly synchronized
by Quicken, one can export the statement in excel and then do some basic
formatting to be used by this program. After this it can be fed to this program
which will convert into QIF format that can be easily imported into Quicken
 
@devnscse
08.10.2023

'''
import os
import re
import pandas as pd
import datetime as dt
import configparser as cp
import global_var as gvar

# This will store the mapping given in the configuration file
# for fast matching
mapping_conf = {}
# The QIF format transformed from the mapping will be added in the list
# at the end it will be written in the file
output_list = []

# Constants that can be changed
# This is the read format of transaction date by the excel reader library
XLS_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

#
# This method is the main method that initiates the transformation process
# based on the configuration provided in the property file. 
#
def transform_to_qif () :
    # Read the configuration
    read_config()

    # Reading the excel
    # Get the input file path
    input_path = get_input_file('Excel input',['.xls','.xlsx'])

    print("Reading the input file for transformation: " + input_path)
    dfXls = pd.read_excel(input_path)

    # Go through each row of excel and parse against the mapping
    # and then write into the array the transformed entry
    if dfXls is None:
        print(f'Error reading the excel file {input_path}')
    else:
        for row_index in range(len(dfXls)) :
            t_date = dfXls.iloc[row_index, gvar.index_t_date]
            t_desc = dfXls.iloc[row_index, gvar.index_t_desc]
            t_type = dfXls.iloc[row_index, gvar.index_t_type]
            t_amt = dfXls.iloc[row_index, gvar.index_t_amt]
            
            if(t_type == 'Debit'):
                t_amt *= -1
            
            output_list.append(transfrom_date(t_date))
            output_list.append("U"+str(t_amt))
            output_list.append("T"+str(t_amt))
            output_list.append("C*")
            transform_category(t_desc)
            output_list.append('^')
    
    # Write the file output
    write_output_list()

#
# This method will transform the date object into QIF date format
#
def transfrom_date(inDate):
    date_object = dt.datetime.strptime(str(inDate),XLS_DATE_FORMAT)
    dt_day = str(date_object.day)
    if date_object.day < 10 :
        dt_day = " " + str(date_object.day)

    century = (date_object.year//100) * 100 
    return "D" + str(date_object.month) + "/" + str(dt_day) + "'" + str(date_object.year - century)    

#
# This method will search for keywords from the configured list from the transaction description.
# And it will try to determine the Category and Memo (if configured)
#
def transform_category(desc):
    global output_list
    strDesc = str(desc)
    found = False

    for mapkey in mapping_conf:
        if mapkey in strDesc.lower():
            found = True
            keyVal = mapping_conf[mapkey]
            output_list.append("P"+mapkey)
            if str(keyVal).endswith(']'):
                # This means the memo also exists then get memo
                transform_category_with_memo(keyVal)
            else:
                output_list.append("L"+keyVal)
            break
    if not found :
        # No mapping found categorize as misc adding the whole text in payee
        output_list.append("P"+strDesc)
        output_list.append("L"+gvar.default_category)

#
# If the category is found to be configured with Memo then this method will be called.
#
def transform_category_with_memo(desc):
    global output_list
    strDesc = str(desc)
    pattern = r'([^[]+)\[([^]]+)\]'
    match = re.search(pattern, strDesc)
    if match:
        output_list.append("M"+match.group(2).strip())
        output_list.append("L"+match.group(1).strip())

#
# This method will read the configuration file
#
def read_config():
    global mapping_conf
    global output_list       

    config_file_path = get_input_file('Properties',['.properties'])
    print(f'Configuration file: {config_file_path} ')
    
    config = cp.ConfigParser()
    config.read(config_file_path)
    
    # Read and populate general configuration    
    gvar.index_t_date = int(config["general"]["col_no_t_date"]) - 1
    gvar.index_t_desc = int(config["general"]["col_no_desc"]) - 1
    gvar.index_t_type = int(config["general"]["col_no_t_type"]) - 1
    gvar.identifier_t_text = config["general"]["col_t_type_credit_text"]
    gvar.index_t_amt = int(config["general"]["col_t_type_amt"]) - 1
    gvar.default_category = config["general"]["default_category"]
    
    output_list.append('!Type:'+config["general"]["type"])

    # Read the mapping and cache it so that it can be used
    # for comparing and identifying the category and memo    
    for keys in config["mapping"]:
        strIden = keys.split(",")
        strValue = config["mapping"][keys]
        for strkey in strIden:
            mapping_conf[strkey.lower().strip()] = strValue

    print("Configuration read successfully, total keywords categorized : " + str(len(mapping_conf)))        
    

#
# This method takes user input for the desired input or configuration files
# First it will check in current directory and display the user with available options
# however if the user does not like then they may select option to type in whole path
# from another direcctory
#
def get_input_file(file_typ_name,allowed_ext):
    input_path = ''
    # Get input files
    # Get the current working directory
    curr_dir = os.getcwd()

    # List all files in the current directory
    input_files = [f for f in os.listdir(curr_dir) if any(f.endswith(ext) for ext in allowed_ext)]
    manual_input = True
    print("\n")
    if(len(input_files) > 0) :
        print(f"Below {file_typ_name} files found in current directory ::: ")
        i = 0
        for f in input_files:
            print(f'[{i}] {f}')
            i += 1

        indx_input = int(input("Type index number to select this file or non-existing other number to manually enter file: "))
        
        if(indx_input <= len(input_files)-1 and indx_input > -1) :
            manual_input = False
            input_path = os.path.join(curr_dir, input_files[int(indx_input)])

    # If user decided to manually input or if no matching files found in the current directory
    if(manual_input):
        input_path = input(f"Please enter {file_typ_name} file path: ")
    
    return input_path

#
# This method will be used to write the output QIF file
#
def write_output_list():
    output_path = input("Enter output file path or just type enter to write here :")
    if(output_path.strip() == '') :
        output_path = os.path.join(os.getcwd(),'output.QIF')

    print('The output file will be written at ' + output_path)
    with open(output_path,"w") as file:
        for str in output_list:
            file.write(str+"\n")

    print("Output has been written....Program Finished")


# Call this main method to begin transformation
transform_to_qif()
