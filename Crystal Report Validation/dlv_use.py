#!/usr/bin/env python3
# Import libraries
from datetime import datetime as dt
from dateutil.parser import parse
import os
import crystal_variables as var

'''
Process the DLV_Use_Log to create dictionaries for production and closing with
report names as keys and the last used date, user ID, and datetime as values.
'''
# Initialize variables
log_path = var.log_path
#log_path = "C:/Users/umh2699/AppData/Local/Programs/Python/Python310/Scripts/practice_log.txt"
excluded_users = var.excluded_users
excluded_folders = var.excluded_folders

# Last used dictionaries - keys:report names, values:date - user ID, datetime
cl_usage = {}
prod_usage = {}
# List of entries in the DLV Use Log
entries = []

###############################################################################
class UseEntry:
    '''
    This class is for entries in the DLV Use Log.

    Attributes:
        name (str): Name of the report
        folder (str): Report's folder (either production or closing)
        user (str): ID of the user who last used the report
        date (str): Last date the report was used
    '''
    def __init__(self, name, folder, user, date):
        '''
        Constructor for a Report Entry.
        '''        
        self.name = name
        self.folder = folder
        self.user = user
        self.date = date

        
###############################################################################
def format_report_name(folder_name, report_name):
    '''
    Formats the name of a report to include the folder and report name.
        Parameters:
            folder_name (str): the name of the folder
            report_name (str): the name of the report
        Return:
            formatted_name (str): the name formatted
    '''
    # Extract the folder name (exclude the env part)
    split_folder = folder_name.split('-')

    # Remove extra words, convert to uppercase, strip whitespace
    folder_name = split_folder[0].upper().replace('REPORTS', '').strip()

    formatted_name = folder_name + ' - ' + report_name.replace('.rpt', '')

    # Remove extra whitespace
    return formatted_name


###############################################################################
def date_str(date):
    '''
    Get the date as a string from datetime object.
        Parameters:
            date: datetime object
        Return:
            The date string as mm/dd/yyyy
    '''
    return date.date().strftime("%m/%d/%Y")


###############################################################################
def update_dictionary(dictionary, entry):
    '''
    Update a dictionary with an entry if the entry is new or has a more recent
    date than the date already stored in the dictionary.
        Parameters:
            dictionary: a dictionary of report names with dates
            entry: a report entry
        Return:
            dictionary: the updated dictionary
    '''
    # The new value that will be added to the dictionary if necessary
    new_value = (date_str(entry.date) + " - " + entry.user, entry.date)
    
    # Report name isn't a key in the dictionary yet
    if entry.name not in dictionary:
        # Update value
        dictionary[entry.name] = new_value

    else:
        # Get the datetime stored for the report
        stored_date = dictionary[entry.name][1]

        # If the current entry's date is more recent than the stored date
        if entry.date > stored_date:
            dictionary[entry.name] = new_value            

    return dictionary


###############################################################################
# Read entries in the DLV_Use_Log text file and close the file at the end
with open(log_path, 'r') as file:

    # For each entry
    for line in file:
        # Fix line's formatting
        line = line.replace('"', '')        
        line = line.replace('\\', '/')
        line = line.lstrip('/')
        line = line.split(',')

        # Get the entry's user ID and convert to uppercase
        user = line[2].upper()
        # Split the parts of the path into a list
        path = line[0].split('/')
        
        # Production/closing and not in the list of excluded users/folders
        if (("Production" in line[0] or "Closing" in line[0]) and
            user not in excluded_users and path[3] not in excluded_folders):

            # Get the formatted report name using the folder and name
            entry_name = format_report_name(path[3], line[1])
            # Get whether it's production or closing
            folder = path[2].split('-')[2].strip()
            # Read the date string into a datetime object
            date = parse(line[5])
            
            # Create an entry with the report name, folder, user ID, and date
            entries.append(UseEntry(entry_name, folder, user, date))

# Sort by folder, then report name, then date
entries.sort(key = lambda x: (x.folder, x.name, x.date))

# Loop through all the reports
for entry in entries:
    
    # If it's a closing entry
    if entry.folder == "Closing":        
        # Update the closing usage dictionary with the entry
        cl_usage = update_dictionary(cl_usage, entry)
        
    else:
        # Update the production usage dictionary
        prod_usage = update_dictionary(prod_usage, entry)


