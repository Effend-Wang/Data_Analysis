import cpk
import test_coverage

# ----------------------------------------------------------------------------
# This is a function to choose analysis mode & run
def mode():
    print("Please choose your work mode:\n\tCPK Analysis (a)\n\tTest Coverage Analysis (b)\nYour choice is:")
    mode_value=input()
    while(1):
        if(mode_value=="a"):
            cpk.cpk_analysis()
            break
        elif(mode_value=="b"):
            test_coverage.test_coverage_analysis()
            break
        else:
            print("No satisfied mode in this program.\nPlease choose again:")
            mode_value=input()
