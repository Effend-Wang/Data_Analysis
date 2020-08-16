# ----------------------------------------------------------------------------
import xlrd
import os
import shutil

# Import program lib
import log

# ----------------------------------------------------------------------------
# Define file path
pro_path=os.getcwd()
result_path=pro_path+"\Result"

# ----------------------------------------------------------------------------
# This is a function to move result file
def result_file_move(file_name):
    shutil.move(file_name,result_path)
    log.write("info","File Operate - File %s have been moved to %s" %(file_name,result_path))

# ----------------------------------------------------------------------------
# This is a function to read excel file
def file_read():
    
    # Read file to data. Get sheet names and output
    print("Please input the source of file:")
    data_source=input()
    print("Opening file %s ..." %data_source)
    data=xlrd.open_workbook(data_source)
    sheetnames=data.sheet_names()
    print("Found sheets: %s" %sheetnames)

    # Choose sheet by "Data" and output
    print("Please choose sheet:")
    chosen_sheet=input()
    print("Reading sheet %s ..." %chosen_sheet)
    datasheet=data.sheet_by_name(chosen_sheet)
    print("Found sheet: %s" %chosen_sheet)

    # Get numbers of rows and output
    nrows=datasheet.nrows
    print("Sheet %s has %d rows" %(chosen_sheet,nrows))

    # Get numbers of cols and output
    ncols=datasheet.ncols
    print("Sheet %s has %d cols" %(chosen_sheet,ncols))

    # Record information of data file
    log.write("info","File Operate - Data File Info:\nFile: %s\nChosen Sheet: %s" %(data_source,chosen_sheet))

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
# Function returns (row,col) of value
def find_value_by_row(worksheet,nrows,value_name,begin_col,end_col):
    result_control=0
    for i in range(nrows):
        workdata=worksheet.row_values(i,start_colx=begin_col,end_colx=end_col)
        for j in range(len(workdata)):
            if (value_name==str(workdata[j])):
                print("Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                log.write("info","File Operate - Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                result_control=1
                break
        if(result_control==1):
            break
    if (result_control==0):
        print("Cannot find out value %s from sheet!Please check file!" %value_name)
        log.write("error","File Operate - Cannot find out value %s from sheet!" %value_name)
    else:
        return i,j

# ----------------------------------------------------------------------------
# This is a function to get value's position from whole sheet by cols
# Function returns (row,col) of value
def find_value_by_col(worksheet,ncols,value_name,begin_row,end_row):
    result_control=0
    for i in range(ncols):
        workdata=worksheet.col_values(i,start_rowx=begin_row,end_rowx=end_row)
        for j in range(len(workdata)):
            if (value_name==str(workdata[j])):
                print("Found '%s' at row %d, col %d" %(value_name,i+1,j+1))
                log.write("info","File Operate - Found '%s' at row %d, col %d" %(value_name,j+1,i+1))
                result_control=1
                break
        if(result_control==1):
            break
    if (result_control==0):
        print("Cannot find out value %s from sheet!Please check file!" %value_name)
        log.write("error","File Operate - Cannot find out value %s from sheet!" %value_name)
    else:
        return j,i
