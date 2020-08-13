# Data Analysis Program
# Author: Effend Wang
# Version: v0.3

# Import local python lib
import sys
import logging
import time
import datetime

# Import program function
import mode_choose

# ----------------------------------------------------------------------------
# Here are messages of program
Program_Name="Data Analysis Program"
Author="Effend Wang"
Version="v0.3"
Last_Edit="2019/08/19 Monday"
Pro_msg="Update in v0.3: \n1. Optimize source code into modular code.\n2. Optimize standard deviation algorithm to avoid computational error.\n3. Add new data in xls file.\n4. Add module name in log file."
Attention_msg="Attention! This software can only run in Vista, Win7, Win10 system!"
Help_msg="If you need help, please see ReadMe."

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
log_path="RunSteps.log"
logging.basicConfig(level=logging.DEBUG,format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',filename=log_path)

# ----------------------------------------------------------------------------
# Program running steps
print("Program Begin\n%s\nAuthor: %s\nLast Edit: %s\n%s\n%s\n%s\n\nStartTime: %s" %(Program_Name,Author,Last_Edit,Pro_msg,Attention_msg,Help_msg,datetime.datetime.now()))
logging.info("Programe Start.")
try:
    mode_choose.mode()
except Exception as e:
    logging.critical("A Program ERROR Occured: %s\n" %e)
