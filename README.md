Migrate Data from Raw Files (Excel Files) to MySQL
==== 

There are 3 files for migrate data from excel files to MySQL Database
    -   Country_Master
    -   Customer_Master
    -   Product_Ledger

## Functions in Code

- split_excel_files: split file from one file that have many table in one sheet spread to many sheets
    -   parameter
            -   sheet_name(list): list of sheet name (table name)
- read_file_to_df: read excel file and put data to dataframe
    -   parameter
            -   file_name(str): excel file name
            -   sheet_name(str): sheet name (table name)
            -   last_updated_date(str): last_updated_date
            -   column(list): list of columns
- load_to_mysql: load data from dataframe to MySQL Database
    -   parameter
            -   df_name(variable): data frame name
            -   table_name(str): target table name (MySQL)
- change_column_datatype: change column type
    -   parameter
            -   df_name(variable): data frame name
            -   column_name(str): column name
            -   data_type(str): str, 'datetime'
- add_digit_customer: add to 7 digits in column custId on Customer_Master table 

### Step

1. run ./mysql-start.sh for start mysql
2. pip install -r requirements.txt
3. run python3 ExcelFiles_to_MySQL.py