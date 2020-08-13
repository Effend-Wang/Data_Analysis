import xlrd
import logging

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
log_path="RunSteps.log"
logging.basicConfig(level=logging.DEBUG,format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',filename=log_path)

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
    logging.info("Data File Info:\nFile: %s\nChosen Sheet: %s" %(data_source,chosen_sheet))

    return datasheet, nrows, ncols

# ----------------------------------------------------------------------------
# This is a function to get 1st col of excel file
def first_col_row(worksheet,begin_row,begin_col):
	first_col_value=worksheet.col_values(0,start_rowx=0,end_rowx=None)
	first_row_value=worksheet.row_values(0,start_colx=0,end_colx=None)

	return first_col_value,first_row_value

# ----------------------------------------------------------------------------
# This is a function to get each col's data
def para_data(worksheet,col_position,begin_row):

    # Reading resent parametric
    res_para=worksheet.cell_value(0,col_position)
    logging.info("Loading parametric: %s." %res_para)

    # Reading resent col values
    workdata=worksheet.col_values(col_position,start_rowx=begin_row,end_rowx=None)
    logging.info("Data loaded.")

    # Return resent col data
    return workdata

# ----------------------------------------------------------------------------
# This is a function to find out value's row
def value_find_row(worksheet,find_row,begin_col,value_name):
    for i in range(len(find_row)):
        if (find_row[i]==value_name):
            value_row=worksheet.row_values(i,start_colx=begin_col,end_colx=None)
            logging.info("Found value %s in row %d" %(value_name,i))
            break
    return i,value_row

# ----------------------------------------------------------------------------
# This is a function to find out value's col
def value_find_col(worksheet,find_col,begin_row,value_name):
    for i in range(len(find_col)):
        if(find_col[i]==value_name):
            value_col=worksheet.col_values(i,start_rowx=begin_row,end_rowx=None)
            logging.info("Found value %s in col %d" %(value_name,i))
            break
    return i,value_col

# ----------------------------------------------------------------------------
# This is a function to choose data range
def data_range():

    # Need user to input an integer number
    print('Please input the number beginning of row:')
    begin_row=int(input())-1
    print('Please input the number beginning of col:')
    begin_col=int(input())-1
    print("Row begins at %d. Col begins at %d" %(begin_row,begin_col))

    # Record information of data range
    logging.info("Chosen Range:\nRow begins at %d\nCol begins at %d" %(begin_row,begin_col))

    return begin_row, begin_col
