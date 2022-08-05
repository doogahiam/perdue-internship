#!/usr/bin/env python3
# Import helpful libraries
import os
from datetime import datetime as dt
import pandas as pd

# Import mods
import dlv_use as dlv
import crystal_variables as var

'''
Creates a Master Inventory Excel file which contains three sheets:
Inventory - Lists the crystal reports in production and indicates when reports
            are missing or have different dates modified in other environments.
            Additional columns for the last date used and user ID for each
            crystal report in production and closing.
Added Reports - Crystal reports found in other environments that are not in the
                master/production directory.
Extras - Paths to extra folders and files found in all environments.
'''

# Environment paths (production, closing, DEV, TEST, QA)
production_path = var.production_path
envs = [var.closing_path, var.dev_path, var.test_path, var.qa_path,
        var.ssdev2_path]

# Excluded folders (according to spec)
excluded_folders = dlv.excluded_folders

# List that store paths to extra files and folders
extra_files = []
extra_folders = []

# Dictionaries for last used info - keys:report names, values:date - user ID
c_last_used = []
p_last_used = []


###############################################################################
def check_excluded(name):
    '''
    Return if a name should be excluded.
        Parameters:
            name: name of a folder to check
        Return:
            True: the folder should be excluded
            False otherwise
    '''
    # Go through the list of excluded folders
    for folder in excluded_folders:

        # True if excluded
        if folder in name:
            return True
        
    return False


###############################################################################
def process_folder(folder, path, reports):
    '''
    Returns a dictionary of properly formatted report names within given folder
    where the keys are report names and values are the report's date modified.
    Keeps track of anything that is not a crystal report (files and folders).
        Parameters:
            folder: current directory to file through
            path: full path to the directory
            reports: the dictionary of all the reports in production
        Return:
            reports (dict): the dictionary updated with the reports found
    '''
    # Raw list of the directories and files within the folder
    content = os.listdir(path)

    # Go through all the contents of the folder
    for item in content:
        # Path to the file/folder
        full_path = os.path.join(path, item)
        
        # If the item is a file
        if os.path.isfile(full_path):
            
            # If item is a crystal report, format and add to the dictionary
            if item.endswith('.rpt'):

                # Get the datetime modified of the file
                date = os.path.getmtime(full_path)
                # Convert the datetime to a formatted string
                dt_mod = dt.fromtimestamp(date)
                dt_mod = dt_mod.strftime("%m/%d/%Y")

                # Add the key-value pair into the reports dictionary where
                # key - report name, value - date modified
                reports[dlv.format_report_name(folder, item)] = dt_mod

            # If the item is not a .db, add to list of extra files
            elif not item.endswith('.db'):
                extra_files.append(full_path)

        # If it's a directory
        elif os.path.isdir(full_path):
            # Add to list of extra folders
            extra_folders.append(full_path)
            
    return reports


###############################################################################
def process_env(env_path):
    '''
    Returns a dictionary of all the reports in an environment.
        Parameters:
            env_path: path of the environment to process
        Return:
            reports_dict: the dictionary of all reports located in the env
    '''
    reports_dict = {}
    
    # Contents of the environment folder
    env_contents = os.listdir(env_path)

    # Go through all items in environment
    for item in env_contents:

        # Get the full path to the current item
        item_path = env_path + item
        
        # If the item is a directory
        if os.path.isdir(item_path) and not check_excluded(item):
            reports_dict = process_folder(item, item_path, reports_dict)

    return reports_dict
        

###############################################################################
def make_df(dictionary):
    '''
    Convert a dictonary with varying column lengths to a dataframe.
        Parameters:
            dictionary: represents column titles with lists of rows
        Return:
            The dictionary converted to a dataframe
    '''
    return pd.DataFrame({k:pd.Series(v) for k,v in dictionary.items()})


###############################################################################
def find_diff(list1, list2):
    '''
    Returns the difference between two lists.
        Parameters:
            list1: guide/original list
            list2: list to check for missing items
        Return:
            list of items in list1 that are missing in list2
    '''
    return sorted(list(set(list1).difference(list2)))


###############################################################################
def track_diff(report_dict):
    '''
    Create a list to track reports in master that are missing in an environment.
        Parameters:
            report_dict: dictionary of reports to compare with master
        Return:
            tracking (list): a list where X represents a missing report
    '''
    # Get the list of report names from the dictionary
    report_list = report_dict.keys()
    # Get a list of reports that are in master but missing in the env list
    env_missing = find_diff(master_list, report_list)
    # List of tracking info for the environment comparison
    tracking = []
    
    # Go through the master reports
    for report in master_list:
        
        # If the report is missing
        if report in env_missing:
            tracking.append("Missing")

        # If the dates the reports were modified aren't matching
        elif master_dict[report] != report_dict[report]:
            tracking.append("Wrong Date")
            
        # If everything lines up/matches, add a blank row
        else:
            tracking.append("")

    return tracking


###############################################################################
def compare_env(env_path):
    '''
    Compare the reports in an environment to the master reports list.
        Parameters:
            env_path: path to the environment that is being checked
        Return:
            A list of resulting information for the environment:
                0 - env_missing: reports in master but not in the env
                1 - env_added: reports in the env but not in master
    '''
    # Get dictionary of files in the environment
    env_dict = process_env(env_path)
    # Get the reports in master but not in the env
    env_missing = track_diff(env_dict)
    # Get the reports in the env but not in master
    env_added = find_diff(env_dict.keys(), master_list)

    return (env_missing, env_added)


###############################################################################
def check_dict(report, dictionary, col):
    '''
    Check if a report is a key in a dictionary and add the corresponding value
    to the col list if it is.
        Parameters:
            report: a string report name
            dictionary: dictionary with report names as keys
            col: a list that represents the column that will be in the excel
        Return:
            col (list): the updated col list
    '''
    # If the report has a last used entry in the dictionary
    if report in dictionary:        
        # Add its information to the column
        col.append(dictionary[report][0])
        
    else:
        # Add a blank
        col.append("")

    return col


###############################################################################
def check_envs(master_inventory):
    '''
    Cross reference master with all other environments to create the master
    inventory and lists of added reports.
        Parameters:
            master_inventory: the dictionary to store the results in
        Return:
            master_inventory: updated master inventory dictionary
    '''
    # For each environment
    for env in envs:

        # Get the current environment's name
        env_name = env.split('/')[4].split(' - ')[2]
        # Compare the environment to the master
        env_results = compare_env(env)
        
        # Update the master and added reports with the results
        master_inventory[env_name] = env_results[0]
        added_reports_dict[env_name] = env_results[1]
        
    return master_inventory

        
###############################################################################
""" Create the master inventory and check other environments """
# Dictionary of reports in production, which will be the master
master_dict = process_env(production_path)
# List of master report names
master_list = master_dict.keys()

# Fill the last used lists based on the dlv usage information
for report in master_list:
    c_last_used = check_dict(report, dlv.cl_usage, c_last_used)
    p_last_used = check_dict(report, dlv.prod_usage, p_last_used)

added_reports_dict = {}
# Add the master reports to the inventory dictonary
master_inventory = {'Report Folder - Report Name': master_list}
# Add results of comparing to other environments
master_inventory = check_envs(master_inventory)
# Add last used information
master_inventory['Last Used SSCLOSE Date - User'] = c_last_used
master_inventory['Last Used SSPROD Date - User'] = p_last_used

# Convert dictionaries to dataframes (make_df for columns with diff lengths)
master_df = pd.DataFrame(master_inventory)
added_reports_df = make_df(added_reports_dict)
extras_df = make_df({'Folders': extra_folders, 'Files': extra_files})

# Create a writer for the excel file and add the dataframes as sheets
writer = pd.ExcelWriter('Crystal Reports Inventory.xlsx', engine = 'openpyxl')
master_df.to_excel(writer, 'Inventory', index = False)
added_reports_df.to_excel(writer, 'Added Reports', index = False)
extras_df.to_excel(writer, 'Extras', index = False)

# Save/close the excel sheet (same directory as the script)
writer.save()

