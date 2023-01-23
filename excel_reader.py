import pandas as pd
import openpyxl
import os


def save_statement_to_output_folder(statement_file, sheet_name, file_name, folder_name):
    file_name = file_name + ".xlsx"
    writer = pd.ExcelWriter(os.path.join("C:/Dev/Tony/Actual Projects/Excel Reader/OutputFile.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='new')
    statement.iloc[statements_to_run].to_excel(writer, sheet_name=sheet_name, index=True)
    
    writer.close()

def try_row(row):    
    statement_file = statement.loc[row].values[1]
    clients = statement.loc[row].values[2]
    sheet_name = statement.loc[row].values[3]
    file_name = statement.loc[row].values[4]
    folder_name = statement.loc[row].values[5]
    print(f"Statement File: {statement_file}\nSheet Name: {sheet_name}\nFolder: {folder_name}\nFile Name: {file_name}")
    save_statement_to_output_folder(statement_file, sheet_name, file_name, folder_name)
    
# Reads in the excel file that I hand created
statements = pd.read_excel("statements.xlsx")

# turns the excel file into a dataframe
statement = pd.DataFrame(statements)

# I imagine this will eventually be handled by an api call or a click in the maintenance window
statements_to_run = input("Which statement would you like to run?")

if statements_to_run == "all":
    for row in range(0, statement['Statement'].count()):
        try_row(row)
else:
    #This is the row where the statement is in the excel file
    statements_to_run = int(statements_to_run) - 2
    try_row(statements_to_run)



