import file_operation
import math
import logging
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Border,Side,Alignment

# ----------------------------------------------------------------------------
# Parameter definition
upp_limit_name="Upper Limit"
low_limit_name="Lower Limit"

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
log_path="RunSteps.log"
logging.basicConfig(level=logging.DEBUG,format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',filename=log_path)

# ----------------------------------------------------------------------------
# This is a function to calculate data average & standard deviaion
def data_cpk_cal(workdata,upp_limit_row,low_limit_row,limit_col_num,begin_col):

    workdata=cpk_data_check(workdata)

    # Calculating max & min of data
    single_max=max(workdata)
    single_min=min(workdata)

    # Calculating average of data
    ndata=len(workdata)
    single_avg=sum(workdata)/ndata
    
    # Calculating standard deviaion of data
    single_std=0.0
    for i in range(ndata):
    	single_std=single_std+math.pow(workdata[i]-single_avg,2)
    single_std=math.pow(single_std/(ndata-1),0.5)
    
    # Calculating CPL, CPU, CPK
    upp_limit=upp_limit_row[limit_col_num-begin_col]
    low_limit=low_limit_row[limit_col_num-begin_col]
    if(upp_limit=="NA" and low_limit=="NA"):
        logging.info("The limit is [NA,NA]")
    elif(upp_limit=="NA"):
        logging.info("The limit is [%s,NA]" %low_limit)
    elif(low_limit=="NA"):
        logging.info("The limit is [NA,%s]" %upp_limit)
    else:
        logging.info("The limit is [%s,%s]" %(low_limit,upp_limit))
    if(single_std==0 or (upp_limit=="NA" and low_limit=="NA")):
    	logging.warning("This parametric std is 0 or limit is NA.")
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

    return single_avg,single_std,single_max,single_min,single_cpl,single_cpu,single_cpk

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
# This is a function to check data, delete blank data
def cpk_data_check(workdata):
    check_data=list(workdata)
    while '' in check_data:
        check_data.remove('')
    new_data=tuple(check_data)
    
    return new_data

# ----------------------------------------------------------------------------
# This is a function to output CPK results to excel file
def output_cpk_result(para,avg,std,pos3_delta,neg3_delta,data_min,data_max,cpl,cpu,cpk,low_limit_row,upp_limit_row):

    # Initial file
    outdata=Workbook()
    outsheet=outdata.worksheets[0]
    outsheet.title='CPK_Result'
    
    # Set style of output data
    # Font Type: Times New Roman
    # Size: 12
    # Border: Thin
    # Alignment: Center
    # Pattern Fill: Red background color when CPK<1.33
    font=Font('Times New Roman',size=12)
    fill=PatternFill(fill_type='solid',start_color='FF0000')
    border=Border(left=Side(border_style='thin',color='000000'),right=Side(border_style='thin',color='000000'),top=Side(border_style='thin',color='000000'),bottom=Side(border_style='thin',color='000000'))
    align=Alignment(horizontal='center',vertical='center')

    # Write excel file
    words=('Parameter','Upper Limit','Lower Limit','AVG','STD','+3Del','-3Del','Max','Min','CPL','CPU','CPK')
    npara=len(avg)
    nwords=len(words)
    for i in range(nwords):
        outsheet.cell(row=i+1,column=1).value=words[i]
    for i in range(npara):
        outsheet.cell(row=1,column=i+2).value=para[i]
        outsheet.cell(row=2,column=i+2).value=upp_limit_row[i]
        outsheet.cell(row=3,column=i+2).value=low_limit_row[i]
        outsheet.cell(row=4,column=i+2).value=avg[i]
        outsheet.cell(row=5,column=i+2).value=std[i]
        outsheet.cell(row=6,column=i+2).value=pos3_delta[i]
        outsheet.cell(row=7,column=i+2).value=neg3_delta[i]
        outsheet.cell(row=8,column=i+2).value=data_max[i]
        outsheet.cell(row=9,column=i+2).value=data_min[i]
        outsheet.cell(row=10,column=i+2).value=cpl[i]
        outsheet.cell(row=11,column=i+2).value=cpu[i]
        outsheet.cell(row=12,column=i+2).value=cpk[i]
        if(cpk[i]!="NA" and cpk[i]<1.33):
                outsheet.cell(row=12,column=i+2).fill=fill
    # Apply style for sheet
    for i in range(nwords):
        for j in range(len(cpk)+1):
            outsheet.cell(row=i+1,column=j+1).font=font
            outsheet.cell(row=i+1,column=j+1).border=border
            outsheet.cell(row=i+1,column=j+1).alignment=align

    outdata.save('CPK_Analysis_Result.xlsx')

# ----------------------------------------------------------------------------
# This is a function to analyse CPK
def cpk_analysis():
    print("CPK Analysis:")
    avg=()
    std=()
    pos3_delta=()
    neg3_delta=()
    data_min=()
    data_max=()
    cpl=()
    cpu=()
    cpk=()
    (worksheet,nrows,ncols)=file_operation.file_read()
    (begin_row,begin_col)=file_operation.data_range()
    (first_col_value,first_row_value)=file_operation.first_col_row(worksheet,begin_row,begin_col)
    (upp_row_num,upp_limit_row)=file_operation.value_find_row(worksheet,first_col_value,begin_col,upp_limit_name)
    (low_row_num,low_limit_row)=file_operation.value_find_row(worksheet,first_col_value,begin_col,low_limit_name)
    para=worksheet.row_values(0,start_colx=begin_col,end_colx=None)
    logging.info("CPK Analysis Details:\n=========================")

    # Calculating by col
    print("Start calculating, please wait...")
    for i in range(ncols):
        if(i<begin_col):
            continue
        else:
            workdata=file_operation.para_data(worksheet,i,begin_row)
            (single_avg,single_std,single_max,single_min,single_cpl,single_cpu,single_cpk)=data_cpk_cal(workdata,upp_limit_row,low_limit_row,i,begin_col)
            # Data are saved as a tuple list in case of no other change
            avg=avg+(single_avg,)
            std=std+(single_std,)
            pos3_delta=pos3_delta+(single_avg+single_std*3,)
            neg3_delta=pos3_delta+(single_avg-single_std*3,)
            data_min=data_min+(single_min,)
            data_max=data_max+(single_max,)
            cpl=cpl+(single_cpl,)
            cpu=cpu+(single_cpu,)
            cpk=cpk+(single_cpk,)

    # Print results
    print("All data calculate finished!\nOutput results as excel file...")
    output_cpk_result(para,avg,std,pos3_delta,neg3_delta,data_min,data_max,cpl,cpu,cpk,low_limit_row,upp_limit_row)
    print("Excel output finished!\nProgram finished!\n")
    logging.info("All CPK analysis steps finished!\nProgram Finished!\n\n")
