import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pymysql
from sqlalchemy import create_engine

def split_excel_files(sheet_name):
    # Read excel file
    df = pd.read_excel('Master.xlsx', header=None)

    # Split by blank rows
    df_list = np.split(df, df[df.isnull().all(1)].index)

    # Create new excel to write the dataframes
    writer = pd.ExcelWriter('Master_Data.xlsx', engine='xlsxwriter')
    for i in range(1, len(df_list) + 1):
        print(sheet_name[i-1])
        df_list[i - 1] = df_list[i - 1].dropna(how='all')
        df_list[i - 1].to_excel(writer, sheet_name=sheet_name[i-1], header=None, index=False)

    # Save the excel file
    writer.save()
    
def read_file_to_df(file_name,sheet_name,last_updated_date,column):
    # Let's select Characters sheet and include column labels
    df = pd.read_excel(file_name, sheet_name = sheet_name, header=None)
    df.columns = column
    df.drop(df.index[:3], inplace=True)
    df.reset_index(drop=True, inplace=True)
    df['last_updated_date'] = last_updated_date

    return df    
    
def load_to_mysql(df_name,table_name):

    engine = create_engine("mysql+pymysql://root:p@ssw0rd@localhost:3306/dbo")
    con = engine.connect()

    df_name.to_sql(name=table_name, con=engine, if_exists = 'replace', index=False)
    
def change_column_datatype(df_name,column_name,data_type):
    
    if (data_type == str):
        df_name[column_name] = df_name[column_name].astype(data_type)
    elif (data_type == 'datetime'):
        df_name[column_name] = pd.to_datetime(df_name[column_name])
    
def add_digit_customer():
    Customer_Master['custId'] = Customer_Master['custId'].apply(lambda x: '{0:0>7}'.format(x)) 

################################################################################################
#--------------------------------------------- Main --------------------------------------------
################################################################################################

#list_sheet    
sheet_name = ['Country_Master','Customer_Master','Product_Ledger']

#split into multiple sheet
split_excel_files(sheet_name)

#list column in each table
Country_Master_columns = ["Country", "Group"]
Customer_Master_columns = ["custId", "custName", "salesChannel", "custCountry"]
Product_Ledger_columns = ["productId", "productName", "stdCost", "stdPrice"]

#read file to dataframe
Country_Master = read_file_to_df('Master_Data.xlsx','Country_Master','20190421', Country_Master_columns)
Customer_Master = read_file_to_df('Master_Data.xlsx','Customer_Master','20190421', Customer_Master_columns)
Product_Ledger = read_file_to_df('Master_Data.xlsx','Product_Ledger','20190421', Product_Ledger_columns)

#change column type of custId from Customer_Master table 
change_column_datatype(Customer_Master,'custId',str)
#change column type of last_updated_date in each table 
change_column_datatype(Customer_Master,'last_updated_date','datetime')
change_column_datatype(Country_Master,'last_updated_date','datetime')
change_column_datatype(Product_Ledger,'last_updated_date','datetime')
#add figit to 7 digits
add_digit_customer()
#load data to MySQL
load_to_mysql(Country_Master,'country_master')
load_to_mysql(Customer_Master,'customer_master')
load_to_mysql(Product_Ledger,'product_ledger')