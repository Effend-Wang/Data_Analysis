# ----------------------------------------------------------------------------
# This .py file is specially used for CPK Analysis
# Main Function: Draw statistic image of data and output as .png file

# Import local python lib
from matplotlib import pyplot
from matplotlib import mlab
import numpy as np

# Import program lib
import file_operation
import log.log as log
import path_config

# ----------------------------------------------------------------------------
# This is a function to check data
def cpkimg_data_check(workdata):
    
    # Data check
    log.write("info","CPK Image - Start checking CPK image data.")
    check_data=list(workdata)
    while '' in check_data:
        check_data.remove('')
    new_data=tuple(check_data)
    log.write("info","CPK Image - Data check finished.")
    return new_data

# ----------------------------------------------------------------------------
# This is a function to create PDF function
def normfun(x,mu,sigma):
    pdf=np.exp(-((x-mu)**2)/(2*sigma**2))/(sigma*np.sqrt(2*np.pi))
    return pdf

# ----------------------------------------------------------------------------
# This is main function to analyse data of CPK
def cpkimg_draw(para_name,data,low_limit,upp_limit,avg,std):

    log.write("info","CPK Image - Start draw CPK distribution image for %s" %para_name)

    # Check data
    workdata=cpkimg_data_check(data)
    title_name=para_name[0]

    # Define x-axis
    x_step_num=20
    if(upp_limit=="NA"):
        true_upp_limit=max(workdata)
    else:
        true_upp_limit=upp_limit
    if(low_limit=="NA"):
        true_low_limit=min(workdata)
    else:
        true_low_limit=low_limit
    x_axis=np.linspace(true_low_limit,true_upp_limit,x_step_num)

    # Initialize image style
    bins=x_axis
    alpha=0.5
    facecolor='gray'
    edgecolor="black"
    #weights=np.ones_like(x_axis)/(len(x_axis))
    
    # Draw frequency histogram
    frequency_each,bins_list,patches=pyplot.hist(workdata,bins=bins,facecolor=facecolor,edgecolor=edgecolor,alpha=alpha)
    pyplot.title(title_name)
    pyplot.xlabel("Value")
    pyplot.ylabel("Quantity")
    pyplot.grid(True)
    
    # Draw upper limit & lower limit & average
    if(low_limit!="NA" and upp_limit!="NA"):
        pyplot.axvline(x=low_limit,ls="-",c="red",label="Limit")
        pyplot.axvline(x=upp_limit,ls="-",c="red")
    elif(upp_limit!="NA"):
        pyplot.axvline(x=upp_limit,ls="-",c="red",label="Limit")
    elif(low_limit!="NA"):
        pyplot.axvline(x=low_limit,ls="-",c="red",label="Limit")
    pyplot.axvline(x=avg,ls="-",c="skyblue",label="Average")
    
    # Draw normal distribution under ideal conditions
    mu=avg
    sigma=std
    if(std=="not use now"):
        norm_x=np.linspace(mu-3*sigma,mu+3*sigma,100)
        pdf=normfun(norm_x,mu,sigma)
        pyplot.plot(norm_x,pdf,ls="-",c="gray",label="Distribution")

    # Draw +3sigma & -3sigma limits
    if(sigma!=0):
        pyplot.axvline(x=mu-3*sigma,ls="--",c="skyblue",label=r'$\pm$3$\sigma$')
        pyplot.axvline(x=mu+3*sigma,ls="--",c="skyblue")

    # Show image
    pyplot.legend()
    
    # Save img
    img_file=title_name+".png"
    pyplot.savefig(img_file)
    result_path=path_config.get_result_path()
    file_operation.result_file_move(img_file,result_path)

    # Close pyplot after image saved
    pyplot.close()

    log.write("info","CPK Image - %s distribution image draw finished.\n=========================" %para_name)
