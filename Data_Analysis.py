# Data Analysis Program
# Author: Effend Wang
# Version: v0.6

# Import local python lib
import sys
import time
import datetime
import os
import shutil

# Import program lib
import mode_choose
import log

# ----------------------------------------------------------------------------
# Here are messages of program
Program_Name="Data Analysis Program"
Author="Effend Wang"
Version="v0.6"
Last_Edit="2019/12/15 Sunday"
Use_msg="This software is still under testing. Not final version."
Attention_msg="Attention! This software can only run in Vista, Win7, Win10 system!"
Help_msg="If you need help, please see ReadMe."

# ----------------------------------------------------------------------------
# Program running steps
print("Program Begin\n%s\nAuthor: %s\nLast Edit: %s\n\n%s\n%s\n%s\n\nStartTime: %s" %(Program_Name,Author,Last_Edit,Use_msg,Attention_msg,Help_msg,datetime.datetime.now()))
log.write("info","Main Func - Program Start.")

# ----------------------------------------------------------------------------
# Set program config
pro_path=os.getcwd()
result_path=pro_path+"\Result"
if os.path.exists(result_path):
    log.write("info","Main Func - \"Result\" dir is exist. Rebuild result dir.")
    shutil.rmtree(result_path)
    os.mkdir(result_path)
    pass
else:
    log.write("info","Main Func - \"Result\" dir is not exist. Build result dir.")
    os.mkdir(result_path)

# Program Begin
try:
    mode_choose.mode()
    os.system('pause')
except Exception as e:
    log.write("critical","A Program ERROR Occured: %s\n" %e)
    print("ERROR: A Program ERROR Occured, check the log file to check the details!")
    os.system('pause')