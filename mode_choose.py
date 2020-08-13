import cpk

# ----------------------------------------------------------------------------
# This is a function to choose analysis mode & run
def mode():
    print("Please choose your work mode:\n\tCPK Analysis (a)\n\tTest Coverage Analysis (b)\nYour choice is:")
    mode_value=input()
    while(1):
        if(mode_value=="a"):
            cpk.cpk_analysis()
            break;
        elif(mode_value=="b"):
            print("Test Coverage Analysis mode cannot use for now.\nPlease choose other mode:")
            mode_value=input()
        else:
            print("No satisfied mode in this program.\nPlease re-choose:")
            mode_value=input()
