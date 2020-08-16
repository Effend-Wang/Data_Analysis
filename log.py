# This file is specific used to writting log

# Import local python lib
import logging

# Import program lib

# ----------------------------------------------------------------------------
# Set logging config
# Logging level includes: logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
log_path="RunSteps.log"
logging.basicConfig(level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s',filename=log_path)

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
        logging.error("Logging write error, please check the code.")
