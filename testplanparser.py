import pandas as pd
import copy
import csv
import sys
import re
import os

def read_test_plan(filename):

    '''
        Function: read_test_plan() function is called by the zephyr_translation() function.
        Argument 1: Filename of the Excel file.
    '''

    # Get the Excel file
    read_file = pd.ExcelFile(filename)

    # Read the sheet names into a list object - remove defaults & "Definitions" sheet.
    sheet_list = read_file.sheet_names
    sheet_list = [x for x in sheet_list if not x.startswith("Sheet") and "Definitions" not in x]

    # This will contain the dicts of TP sheets. Type: <list> of <dicts>
    excel_dicts = []

    for member in sheet_list:
        # Parse applicable sheets and add to list
        excel_dicts.append(read_file.parse(member, skiprows=1).to_dict())

    return(excel_dicts, sheet_list)

def zephyr_translation(test_plan_excel, jira_users = None):

    '''
        Function: read_test_plan() function is called by the zephyr_translation() function.
        Argument 1: Filename of the Excel file.
    '''


    # Columns for the Zephyr import CSV
    headers = ["Labels", "Name", "Objective", "Owner", "Priority", "Status", "Estimated Time"]

    # Time estimate - default of 8 hours. Required Format - "hh:mm"
    default_estimate = "08:00"
    # Status - default of Not Started. Required Format - "Draft"
    status = "Draft"
    # Priority Status - defaults to normal
    default_priority = "Normal"

    # Populate the user field

    if jira_users:
        default_user = [jira_users[i][1] for i in range(len(jira_users)) if "Sarmad" in jira_users[i][0]]
        default_user=''.join(default_user)
    else:
        default_user = "Unassigned"

    # Call the testplan reader function
    TP_sheets, TP_sheet_names = read_test_plan(test_plan_excel)

    # Iterate loop counter
    iter = 0

    # Initialise test plan folder
    dir_path = (sys.argv[1] + "\\" + sys.argv[2].replace(".xlsx", ""))

    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


    # Iterate through the list - type: Dict of dicts
    for sheet in TP_sheets:

        # Copy dict of dict to avoid changing iterating target
        copy_sheet = copy.copy(sheet)

        # Column Loop Initialisers
        copy_sheet[headers[0]] = {}
        copy_sheet[headers[1]] = {}
        test_objectives_list = []
        test_count = 0

        TP_columns = []
        [TP_columns.append(columns) for columns in list(sheet.keys()) if "Unnamed" not in columns]

        # Add the label & name columns
        for j in copy_sheet[TP_columns[4]]:
            chip_id = copy_sheet[TP_columns[1]][j]
            volt = copy_sheet[TP_columns[4]][j]
            temp = copy_sheet[TP_columns[5]][j]

            # Regex to remove Sheet names/Test Types with numbers & non-underscore characters in them
            TP_sheet_names[iter] = re.sub(r'[^A-Za-z_]+', "", TP_sheet_names[iter])
            test_type = TP_sheet_names[iter]

            # Populate Test Case No & Label
            copy_sheet[headers[0]][j] = copy_sheet[TP_columns[3]][j]
            copy_sheet[headers[1]][j] =  test_type + str(copy_sheet[TP_columns[0]][j]) + "_" + chip_id + "_"\
                                         + volt + "_"+ temp

            # Populate string of independent variables
            if len(TP_columns) > 6:
                frequency = copy_sheet[TP_columns[6]][j]
                vars_string = (f'over frequency range {frequency} at ({temp}) degrees C and ({volt})V')
            else:
                vars_string = (f'at ({temp}) degrees C and ({volt})V')

            # Write the objective of the test.
            test_objectives_list.append(f'Measure {test_type} of the {chip_id} {vars_string}')

            test_count+=1

        # Write to the Zephyr import CSV
        with open(dir_path + "\\" +  TP_sheet_names[iter] + ".csv", 'w', newline = '') as csvfile:
             writer = csv.writer(csvfile)
             writer.writerow(headers)
             for i in range(test_count):
                writer.writerow([list(copy_sheet[headers[0]].values())[i], list(copy_sheet[headers[1]].values())[i]
                                    , test_objectives_list[i], default_user, default_priority, status, default_estimate])

        # Delete the previous TestPlan columns to prepare for next loop cycle
        [copy_sheet.pop(TP_columns[column]) for column in range(len(TP_columns))]

        # Iterate for next loop cycle
        iter+=1


def RFIT_jira(export_csv_filename):

    # Read CSV with JIRA user info
    jira_csv = pd.read_csv(export_csv_filename)

    #Populate dict of usernames & IDs
    jira_dict = dict(zip(jira_csv["User name"].tolist(),jira_csv["User id"].tolist()))
    return(jira_dict)

def write_dict_to_csv(filename, dict_input):

    # Writes each key:value pair of dict input into the csv file input
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            [writer.writerow(s) for s in dict_input]
            print("\nWritten Dictionary to CSV")

def print_helper():

    print(    '''
        #################
        Testplanparser.py
        #################
        
        Argument 1: Workspace Directory
        Argument 2: Filename of the Test Plan Excel file (xlsx format expected).
    ''')

if __name__ == "__main__":
    print_helper()

    # List of RFIT members
    RFIT = ["Sarmad", "Themba", "Aish", "Sailesh", "Jun", "Hassan", "Vikas"]

    # Populate Jira info for team members in "RFIT" list
    RFIT_Team = [(key, value) for key, value in RFIT_jira(sys.argv[1] + "\JIRA Users Export.csv").items() if
                 any(member in key for member in RFIT)]

    # Write RFIT team JIRA IDs to a CSV
    write_dict_to_csv(sys.argv[1] + "\RFIT.csv", RFIT_Team)

    zephyr_translation(sys.argv[1] + "\\" + sys.argv[2], RFIT_Team)

