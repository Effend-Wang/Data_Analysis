# This file is specific used to writting log

# Import local python lib
import logging

# Import program lib
import path_config

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
def log_style():
    log_path=path_config.get_log_path()
    log_filename=log_path+"\\RunSteps.log"
    logging.basicConfig(level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s',filename=log_filename)

def write(level,log_message):
    if (level=="info"):
        logging.info(log_message)
    elif (level=="warning"):
        logging.warning(log_message)
    elif (level=="error"):
        logging.error(log_message)
    elif (level=="critical"):
        logging.critical(log_message)
    else:
        logging.error("Level %s is not defined/exist. Please check the code." %level)
