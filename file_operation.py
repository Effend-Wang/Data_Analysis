# ----------------------------------------------------------------------------
import xlrd
import os
import shutil

# Import program lib
import log.log as log
import path_config

# ----------------------------------------------------------------------------
# Define file path
pro_path=os.getcwd()

# ----------------------------------------------------------------------------
# This is a function to move result file
def result_file_move(file_name,target_path):

    shutil.move(file_name,target_path)
    log.write("info","File %s have been moved to path %s." %(file_name,target_path))

# ----------------------------------------------------------------------------
# This is a function to read excel file
# Return datasheet,nrows,ncols
def file_read():
    
    # Read excel file and get sheet names.
    print("Input the path of file:")
    data_source=input()
    
    # When draw file (with space in path) into terminal, there will have " charactor in path
    # Delete " charactor if exist
    while '\"' in data_source:
        data_source.remove('\"')

    # Check the file is exist or not
    while (path_config.path_check(data_source)==False):
        print("The file path is not exist! Please input again.\nInput the path of file:")
        data_source=input()
        # Delete "" charactor if exist
        while '\"' in data_source:
            data_source.remove('\"')

    # Loading data file, get data and sheetnames
    print("Loading data from %s.\nPlease wait for a while, it may need some times depend on size of file..." %data_source)
    data=xlrd.open_workbook(data_source)
    sheetnames=data.sheet_names()
    print("File Loading finished. Found sheet: %s" %sheetnames)
    log.write("info","Loading file %s" %data_source)
    log.write("info","File includes sheet %s" %sheetnames)

    # Choose sheet by "Data" and output
    print("Choose the sheet you want to analyse:")
    chosen_sheet=input()
    # Check if the sheet exist
    while chosen_sheet not in sheetnames:
        print("The sheet you chose is not exist. Please input again.\nChoose the sheet you want to analyse:")
        chosen_sheet=input()
    print("Loading sheet %s, please wait..." %chosen_sheet)
    datasheet=data.sheet_by_name(chosen_sheet)
    print("Sheet loading finished.")
    log.write("info","Loading sheet %s finished" %chosen_sheet)

    # Get numbers of rows and output
    nrows=datasheet.nrows
    print("Sheet %s has %d rows." %(chosen_sheet,nrows))
    log.write("info","Sheet %s has %d rows" %(chosen_sheet,nrows))

    # Get numbers of cols and output
    ncols=datasheet.ncols
    print("Sheet %s has %d cols.\n" %(chosen_sheet,ncols))
    log.write("info","Sheet %s has %d cols" %(chosen_sheet,ncols))
    
    return datasheet,nrows,ncols

# ----------------------------------------------------------------------------
# This is a function to get one col data from sheet
# Function returns single col data, with specific begin row and end row.
def one_col_data(worksheet,col_num,begin_row,end_row):

    workdata=worksheet.col_values(col_num,start_rowx=begin_row,end_rowx=end_row)
    
    return workdata

# ----------------------------------------------------------------------------
# This is a function to get one row data from sheet
# Function returns single row data, with specific begin col and end col.
def one_row_data(worksheet,row_num,begin_col,end_col):
    
    workdata=worksheet.row_values(row_num,start_colx=begin_col,end_colx=end_col)
    
    return workdata

# ----------------------------------------------------------------------------
# This is a function to get value's position from whole sheet by rows
# Return (row,col) of worksheet
def find_value_by_row(worksheet,nrows,value_name,begin_col,end_col):
    
    result_control=0
    for i in range(nrows):
        workdata=worksheet.row_values(i,start_colx=begin_col,end_colx=end_col)
        for j in range(len(workdata)):
            if (value_name==str(workdata[j])):
                print("Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                log.write("info","Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                result_control=1
                break
        if(result_control==1):
            break
    if (result_control==0):
        print("Cannot find out value %s from sheet!Please check file!" %value_name)
        log.write("error","Cannot find out value %s from sheet!" %value_name)
    else:
        return i,j

# ----------------------------------------------------------------------------
# This is a function to get value's position from whole sheet by cols
# Return (row,col) of worksheet
def find_value_by_col(worksheet,ncols,value_name,begin_row,end_row):
    
    result_control=0
    for i in range(ncols):
        workdata=worksheet.col_values(i,start_rowx=begin_row,end_rowx=end_row)
        for j in range(len(workdata)):
            if (value_name==str(workdata[j])):
                print("Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                log.write("info","Found '%s' at row %d, col %d" %(value_name,j+1,i+1))
                result_control=1
                break
        if(result_control==1):
            break
    if (result_control==0):
        print("Cannot find out value %s from sheet!Please check file!" %value_name)
        log.write("error","Cannot find out value %s from sheet!" %value_name)
    else:
        return j,i
