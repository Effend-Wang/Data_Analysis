# Data Analysis Program
# Author: Effend Wang
# Version: v0.7

# Import local python dylib
import sys
import datetime
import os
import shutil

# Import program's python file
import mode_choose
import log.log as log
import path_config

# ----------------------------------------------------------------------------
# Setting messages of program
program_name="Data Analysis"
author="Effend Wang"
version="v0.7"
last_edit="2020/08/16"
use_msg="Fixed output bug of test coverage and CPK. Added independent result folder."
attention_msg="Attention! This software can only run in Vista, Win7, Win10 system!"
help_msg="If you need help, please see ReadMe or contact with developer effend_wang@outlook.com"

# ----------------------------------------------------------------------------
# Get the running time of program
# Time style: YearMonthDay Hour-Minute-Second
program_start_time=datetime.datetime.now().strftime("%Y%m%d %H-%M-%S")

# ----------------------------------------------------------------------------
# Set path config
# Result path style: (program path)\Result (program start time)\
# Log path style: (result path)\log

# Create the result path
path_config.result_path(program_start_time)
# Create the log path
path_config.log_path()
# Get result path
result_path=path_config.get_result_path()
# Get log path and setting log style
log_path=path_config.get_log_path()
log.log_style()

# ----------------------------------------------------------------------------
# Output program messages

# Output messages in log
log.write("info","\n%s %s" %(program_name,version))
log.write("info","%s" %last_edit)
log.write("info","Author: %s" %author)
log.write("info","Program Start!\n")

# Output messages in terminal
print('*'*50)
print("%s %s\nAuthor: %s\n%s\n%s\n" %(program_name,version,author,attention_msg,help_msg))
print("Program Start Time: %s" %program_start_time)
print("Set Result Path: %s" %result_path)
print("Set Log Path: %s\n" %log_path)
print('*'*50+"\n")

# ----------------------------------------------------------------------------
# Program start
#try:
	# Choose analyse mode
mode_choose.mode()
log.write("info","Program Running Finished! Program End!\n\n")
print("Program running finished! Please check the result in:\n%s" %result_path)
os.system('pause')
#except Exception as e:
    # Record and output exception message
    #log.write("warning","Program Running Fail!")
    #log.write("critical","A Program ERROR Occured: %s\n\n" %e)
    #print("ERROR: A Program ERROR Occured!\n%s" %e)
    #os.system('pause')