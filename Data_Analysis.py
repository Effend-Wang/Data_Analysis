# Data Analysis Program
# Author: Effend Wang
# Version: v0.2

# Import math for calculating data
import math

# Import xlrd for reading excel file
import xlrd

# Import xlwt for writting results to excel file
import xlwt

# Import sys for system operation
import sys

# Import logging for writting output to log file
import logging

# Import time for getting resent time
import time
import datetime

# ----------------------------------------------------------------------------
# Here are messages of program
Program_Name="Data Analysis Program"
Author="Effend Wang"
Version="v0.2"
Last_Edit="2019/08/15 Thursday"
Pro_msg="In this version, you can input .xlsx or .xls file to program. Program will output CPK result file as .xls file, also will output a .log file for debug."
Attention_msg="Attention! You need first delete Analysis_Result.xls file if it exist in program's path!"
Help_msg="If you need help, please see ReadMe."

# ----------------------------------------------------------------------------
# Here are parameter need use in program
upp_limit_name="Upper Limit"
low_limit_name="Lower Limit"
log_path="RunSteps.log"

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
logging.basicConfig(level=logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s',filename=log_path)

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
    logging.info("Loading parametric data: %s." %res_para)

    # Reading resent col values
    workdata=worksheet.col_values(col_position,start_rowx=begin_row,end_rowx=None)
    logging.info("Data loaded.")

    # Return resent col data
    return workdata

# ----------------------------------------------------------------------------
# This is a function to calculate data average & standard deviaion
def data_cpk(workdata,upp_limit_row,low_limit_row,limit_col_num,begin_col):

    # Calculating average of data
    ndata=len(workdata)
    single_avg=sum(workdata)/ndata
    
    # Calculating standard deviaion of data
    single_std=0
    for i in range(ndata):
    	single_std=single_std+math.pow(workdata[i]-single_avg,2)
    single_std=math.sqrt(single_std/ndata)
    
    # Calculating CPL, CPU, CPK
    upp_limit=upp_limit_row[limit_col_num-begin_col]
    low_limit=low_limit_row[limit_col_num-begin_col]
    if(upp_limit=="NA" and low_limit=="NA"):
        logging.info("The limit is [NA,NA]")
    elif(upp_limit=="NA"):
        logging.info("The limit is [%f,NA]" %low_limit)
    elif(low_limit=="NA"):
        logging.info("The limit is [NA,%f]" %upp_limit)
    else:
        logging.info("The limit is [%s,%s]" %(low_limit,upp_limit))
    if(single_std==0 or (upp_limit=="NA" and low_limit=="NA")):
    	logging.warning("This parametric std or limit is NA.")
    	single_cpl="NA"
    	single_cpu="NA"
    	single_cpk="NA"
    elif(upp_limit=="NA"):
    	logging.warning("This parametric has no upper limit.")
    	single_cpl=(single_avg-low_limit)/3/single_std
    	single_cpu="NA"
    	single_cpk=single_cpl
    elif(low_limit=="NA"):
    	logging.warning("This parametric has no lower limit.")
    	single_cpl="NA"
    	single_cpu=(upp_limit-single_avg)/3/single_std
    	single_cpk=single_cpu
    else:
    	single_cpl=(single_avg-low_limit)/3/single_std
    	single_cpu=(upp_limit-single_avg)/3/single_std
    	single_cpk=min(single_cpl,single_cpu)

    # Output calculating results
    logging.info("The STD of this parametric is %f",single_std)
    logging.info("The AVG of this parametric is %f",single_avg)
    if(single_cpk!="NA"):
        logging.info("The CPK of this parametric is %f",single_cpk)
    else:
        logging.info("The CPK of this parametric is NA")

    # Check values
    check_value(single_cpk)

    return single_avg,single_std,single_cpl,single_cpu,single_cpk

# ----------------------------------------------------------------------------
# This is a function to check CPK values
def check_value(cpk_value):
    if(cpk_value=="NA"):
        logging.error("Parametric has no CPK value.\n=========================")
    elif(cpk_value<1.33):
        logging.error("Parametric CPK is not OK.\n=========================")
    else:
        logging.info("Parametric CPK is OK.\n=========================")

# ----------------------------------------------------------------------------
# This is a function to find out value's row
def value_find_row(worksheet,find_row,begin_col,value_name):
    for i in range(len(find_row)):
        if (find_row[i]==value_name):
            value_row=worksheet.row_values(i,start_colx=begin_col,end_colx=None)
            logging.info("Found value %s in row %d" %(value_name,i))
            break
    return value_row

# ----------------------------------------------------------------------------
# This is a function to find out value's col
def value_find_col(worksheet,find_col,begin_row,value_name):
    for i in range(len(find_col)):
        if(find_col[i]==value_name):
            value_col=worksheet.col_values(i,start_rowx=begin_row,end_rowx=None)
            logging.info("Found value %s in col %d" %(value_name,i))
            break
    return value_col

# ----------------------------------------------------------------------------
# This is a function to output CPK results to excel file
def output_cpk_result(para,avg,std,cpl,cpu,cpk,low_limit_row,upp_limit_row):

    # Initial file
    outdata=xlwt.Workbook(encoding='ascii')
    outsheet=outdata.add_sheet('Result')

    # Set style of output data
    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Times New Roman'
    font.bold=False
    font.underline=False
    font.italic=False
    style.font=font

    # Write excel file
    outsheet.write(0,0,'Parametric',style)
    outsheet.write(1,0,'Upper Limit',style)
    outsheet.write(2,0,'Lower Limit',style)
    outsheet.write(3,0,'AVG',style)
    outsheet.write(4,0,'STD',style)
    outsheet.write(5,0,'CPL',style)
    outsheet.write(6,0,'CPU',style)
    outsheet.write(7,0,'CPK',style)
    npara=len(avg)
    for i in range(npara):
        outsheet.write(0,i+1,para[i],style)
        outsheet.write(1,i+1,upp_limit_row[i],style)
        outsheet.write(2,i+1,low_limit_row[i],style)
        outsheet.write(3,i+1,avg[i],style)
        outsheet.write(4,i+1,std[i],style)
        outsheet.write(5,i+1,cpl[i],style)
        outsheet.write(6,i+1,cpu[i],style)
        outsheet.write(7,i+1,cpk[i],style)

    outdata.save('CPK_Analysis_Result.xls')

# ----------------------------------------------------------------------------
# This is a function to analyse CPK
def cpk_analysis():
    print("CPK Analysis:")
    avg=()
    std=()
    cpl=()
    cpu=()
    cpk=()
    (worksheet,nrows,ncols)=file_read()
    (begin_row,begin_col)=data_range()
    (first_col_value,first_row_value)=first_col_row(worksheet,begin_row,begin_col)
    upp_limit_row=value_find_row(worksheet,first_col_value,begin_col,upp_limit_name)
    low_limit_row=value_find_row(worksheet,first_col_value,begin_col,low_limit_name)
    para=worksheet.row_values(0,start_colx=begin_col,end_colx=None)
    logging.info("CPK Analysis Details:\n=========================")

    # Calculating by col
    print("Start calculating, please wait...")
    for i in range(ncols):
        if(i<begin_col):
            continue
        else:
            workdata=para_data(worksheet,i,begin_row)
            (single_avg,single_std,single_cpl,single_cpu,single_cpk)=data_cpk(workdata,upp_limit_row,low_limit_row,i,begin_col)
            # Data are saved as a tuple list in case of no other change
            avg=avg+(single_avg,)
            std=std+(single_std,)
            cpl=cpl+(single_cpl,)
            cpu=cpu+(single_cpu,)
            cpk=cpk+(single_cpk,)

    # Print results
    print("All data calculate finished!\nOutput results as excel file...")
    output_cpk_result(para,avg,std,cpl,cpu,cpk,low_limit_row,upp_limit_row)
    print("Excel output finished!\nProgram finished!\n")
    logging.info("All steps finished!\nProgram Finished!\n\n")

# ----------------------------------------------------------------------------
# This is a function to analyse Test Coverage
def test_coverage_analysis():
    print("Test Coverage Analysis:")
    (worksheet,nrows,ncols)=file_read()

# ----------------------------------------------------------------------------
# This is a function to choose analysis mode & run
def mode():
    print("Please choose your work mode:\n\tCPK Analysis (a)\n\tTest Coverage Analysis (b)\nYour choice is:")
    mode_value=input()
    while(1):
        if(mode_value=="a"):
            cpk_analysis()
            break;
        elif(mode_value=="b"):
            print("Test Coverage Analysis mode cannot use for now.\nPlease choose other mode:")
            mode_value=input()
        else:
            print("No satisfied mode in this program.\nPlease re-choose:")
            mode_value=input()

# ----------------------------------------------------------------------------
# Program running steps
print("Program Begin\n%s\nAuthor: %s\nLast Edit: %s\n%s\n%s\n%s\n\nRunTime: %s" %(Program_Name,Author,Last_Edit,Pro_msg,Attention_msg,Help_msg,datetime.datetime.now()))
logging.info("Programe Start.")
try:
    mode()
except Exception as e:
    logging.critical("A Program ERROR Occured: %s\n" %e)
