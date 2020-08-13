# Data Analysis Program
# Author: Effend Wang
# Version: 0.1

# Attention!In this version, you can only calculating average and standard deviation of data.

# Import math for futher analysis, need use --hiden-import math when package
import math

# Import xlrd for reading excel file, need use --hiden-import xlrd when package
import xlrd

# Import xlwt for writting results to excel file, need use --hiden-import xlwt when package
import xlwt

# This is a function to read excel file
def file_read():
    # Read file to data. Get sheet names and output(need an if)
    print("Please input the source of file:")
    data_source=input()
    print("Opening file...")
    data=xlrd.open_workbook(data_source)
    sheetnames=data.sheet_names()
    print("Found sheets: %s" %sheetnames)

    # Choose sheet by "Data" and output(need an if)
    print("Please choose sheet:")
    chosen_sheet=input()
    print("Reading sheet...")
    datasheet=data.sheet_by_name(chosen_sheet)
    print("Found sheet: %s" %chosen_sheet)

    # Get numbers of rows and output
    nrows=datasheet.nrows
    print("Sheet %s has %d rows" %(chosen_sheet,nrows))

    # Get numbers of cols and output
    ncols=datasheet.ncols
    print("Sheet %s has %d cols" %(chosen_sheet,ncols))

    return datasheet, nrows, ncols

# This is a function to choose data range
def data_range():
    print('Please input the number beginning of row:')
    begin_row=int(input())
    print('Please input the number beginning of col:')
    begin_col=int(input())
    print("Row begins at %d. Col begins at %d" %(begin_row,begin_col))

    return begin_row, begin_col

# This is a function to get each parametric's data
def para_data(worksheet,i,begin_row):
    print("Loading data...")
    workdata=worksheet.col_values(i,start_rowx=begin_row,end_rowx=None)
    print("Data loaded.")
    
    return workdata

# This is a function to calculate data average & standard deviation
def data_avg_std(workdata):
    # Calculating average of data
    ndata=len(workdata)
    single_avg=sum(workdata)/ndata
    # Calculating standard deviation of data
    single_std=0
    for i in range(ndata):
        single_std=single_std+math.pow(workdata[i]-single_avg,2)
    single_std=math.sqrt(single_std/ndata)
    
    return single_avg,single_std

# This is a function to calculate data

# This is a function to output results to excel file
def output_result(avg,std):
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
    outsheet.write(0,0,'AVG',style)
    outsheet.write(1,0,'STD',style)
    len_avg=len(avg)
    len_std=len(std)
    for i in range(len_avg):
        outsheet.write(0,i+1,avg[i],style)
    for i in range(len_std):
        outsheet.write(1,i+1,std[i],style)
    outdata.save('Analysis_Result.xls')

# Main steps
print("Data Analysis Program v0.1\nAuthor: Effend Wang\nAttention! In this version, you can only calculating average and standard deviation of data.")
(worksheet,nrows,ncols)=file_read()
(begin_row,begin_col)=data_range()
avg=()
std=()
# Calculating by col
print("Start calculating...")
for i in range(ncols):
    if i<begin_col:
        continue
    else:
        workdata=para_data(worksheet,i,begin_row)
        (single_avg,single_std)=data_avg_std(workdata)
        # Data are saved as a tuple list in case of no other change
        avg=avg+(single_avg,)
        std=std+(single_std,)
# Print results
print("All data calculate finished!\nOutput excel file...")
output_result(avg,std)
print("File output finished!")
