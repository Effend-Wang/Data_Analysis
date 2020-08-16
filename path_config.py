# This file is used to:
# 1. Check path
# 2. Create path
# 3. Get path

# Import local python dylib
import os
import shutil
import sys

# Import program's python file
import log.log as log

# ----------------------------------------------------------------------------
# Here is the function to check whether path exist

# Check path exist or not, return Bool variable
def path_check(path):
	if os.path.exists(path):
		return True
	else:
		return False

# ----------------------------------------------------------------------------
# Here are functions to create path

# Create path and return the path
def path_create(create_path):
	if (path_check(create_path)==False):
		os.mkdir(create_path)
	else:
		print("Warning: Path %s is already exist! This may takes error to program." %create_path)

	return create_path

# Create result path
# Include 1 global item result_path_global
# Result Path must be defined first! This is main folder of program!
def result_path(program_start_time):
	pro_path=os.getcwd()
	global result_path_global
	result_path_global=pro_path+"\\Result "+program_start_time
	path_create(result_path_global)

# Create log path
# Include 1 global item log_path_global
def log_path():
	global log_path_global
	log_path_global=result_path_global+"\\log"
	path_create(log_path_global)

# ----------------------------------------------------------------------------
# Here are functions to get path
# Warning: These path parameter should be global defined before!

# Get result path
def get_result_path():
	if (path_check(result_path_global)==False):
		print("Error: The result path is not defined before! Program will report error!")
	
	return result_path_global

# Get log path
def get_log_path():
	if (path_check(log_path_global)==False):
		print("Error: The log path is not defined before! Program will report error!")
	
	return log_path_global

# Get the path of program
def get_program_path():
	
	return os.getcwd()