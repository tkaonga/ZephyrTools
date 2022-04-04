import pandas as pd
import copy
import csv
import sys
import re
import os


def read_test_plan(filename, skip_row):
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
        excel_dicts.append(read_file.parse(member, skiprows=skip_row).to_dict())

    return (excel_dicts, sheet_list)


def zephyr_translation(test_plan_excel, skip_row_len, jira_users=None):
    '''
        Function: read_test_plan() function is called by the zephyr_translation() function.
        Argument 1: Filename of the Excel file.
    '''

    # Columns for the Zephyr import CSV
    headers = ["Labels", "Name", "Objective", "Owner", "Priority", "Status", "Estimated Time", "Folder", "Precondition"]

    # Time estimate - default of 8 hours. Required Format - "hh:mm"
    default_estimate = "08:00"
    # Status - default of Not Started. Required Format - "Draft"
    status = "Draft"
    # Priority Status - defaults to low
    default_priority = "Low"
    # Default Precondition
    default_precondition = "Baseline"

    # Populate the user field

    if jira_users:
        default_user = [jira_users[i][1] for i in range(len(jira_users)) if "Sailesh" in jira_users[i][0]]
        default_user = ''.join(default_user)
    else:
        default_user = "Unassigned"

    # Call the testplan reader function
    TP_sheets, TP_sheet_names = read_test_plan(test_plan_excel, skip_row=skip_row_len)

    # Iterate loop counter
    iter = 0

    # Get TP prefix
    test_plan_prefix = sys.argv[2].replace(".xlsx", "")

    # Initialise test plan folder & filepath
    dir_path = (sys.argv[1] + "\\" + test_plan_prefix)
    file_path = dir_path + "\\" + test_plan_prefix + "_Zephyr_Import" + ".csv"

    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    # Create new CSVfile to reference in loop, if it doesn't exist
    f = open(file_path, 'w', newline='')
    writer = csv.writer(f)
    writer.writerow(headers)
    f.close()
    f = open(file_path, 'a', newline='')

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
            folder = ""

            # Regex to remove Sheet names/Test Types with numbers & non-underscore characters in them
            TP_sheet_names[iter] = re.sub(r'[^A-Za-z_]+', "", TP_sheet_names[iter])
            test_type = TP_sheet_names[iter]

            split_type = test_type.split("_")

            # Parse test cases into Small Signal Gain folder
            if len(split_type) > 1 and "SSG" in split_type[0]:
                folder = f"Small Signal Gain: {split_type[1]}"

            # Populate Test Case No & Label
            copy_sheet[headers[0]][j] = copy_sheet[TP_columns[3]][j]
            copy_sheet[headers[1]][j] = test_type + str(copy_sheet[TP_columns[0]][j]) + "_" + chip_id + "_" \
                                        + volt + "_" + temp

            # Populate string of independent variables
            if "Frequency" in TP_columns:
                frequency = copy_sheet[TP_columns[6]][j]
                vars_string = (f'over frequency range {frequency} at ({temp}) degrees C and ({volt})V')
            else:
                vars_string = (f'at ({temp}) degrees C and ({volt})V')

            # PRIORITY & PRECONDITION: Raise priority to "High" if there is a Mk1 label
            mark_one_state = copy_sheet[TP_columns[7]]
            if isinstance(mark_one_state[j], str) and mark_one_state[j] == 'Y':
                priority = "High"
                precondition = "Mk1"
            else:
                priority = default_priority
                precondition = default_precondition

            # Write the objective of the test.
            test_objectives_list.append(f'Measure {test_type} of the {chip_id} {vars_string}')

            # # Write to the Zephyr import CSV
            writer = csv.writer(f)
            writer.writerow([list(copy_sheet[headers[0]].values())[j], list(copy_sheet[headers[1]].values())[j]
                                , test_objectives_list[j], default_user, priority, status, default_estimate,
                             folder, precondition])

            test_count += 1

        # Delete the previous TestPlan columns to prepare for next loop cycle
        [copy_sheet.pop(TP_columns[column]) for column in range(len(TP_columns))]

        # Iterate for next loop cycle
        iter += 1

    # Closes CSV after loop execution
    f.close()


def RFIT_jira(export_csv_filename):
    # Read CSV with JIRA user info
    jira_csv = pd.read_csv(export_csv_filename)

    # Populate dict of usernames & IDs
    jira_dict = dict(zip(jira_csv["User name"].tolist(), jira_csv["User id"].tolist()))
    return (jira_dict)


def write_dict_to_csv(filename, dict_input):
    # Writes each key:value pair of dict input into the csv file input
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        [writer.writerow(s) for s in dict_input]
        print("\nWritten Dictionary to CSV")


def print_helper():
    print('''
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

    zephyr_translation(sys.argv[1] + "\\" + sys.argv[2], 4, RFIT_Team)

