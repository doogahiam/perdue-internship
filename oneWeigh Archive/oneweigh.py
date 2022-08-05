#!/usr/bin/env python3

'''
Check local machines and archive/remove tickets basesd on set variables.

Required modules: pandas, openpyxl 
'''
import ow_variables as var
from datetime import datetime, timedelta
from dateutil.parser import parse
import logging
import os
import shutil
import pandas as pd

#Make the ages timedelta objects in days
local_max = timedelta(var.local_max * 365) # Max age of local machine files
archive_max = timedelta(var.archive_max * 365) # Max age of archive files

#Import other variables
archive_path = var.archive_path # Path to the archived files
machines_excel = var.machines_excel #Path to excel of the machines info
log_path = var.log_path #Path to log file to write to
log_details = var.log_details # True - show all details, False - show totals

machine_paths = [] # Paths to all the local machines
archive = [] # List of files in the archive
num_archived = 0 # Track files copied to the archive
num_local_removed = 0 # Track old files removed from local machines


def ping(serial):
    '''
    Parameter:
        serial (str): serial of a machine to ping
    Return:
        True if the machine is pinged successfully, False if it is not.
    '''
    return os.system("ping " + serial) == 0

def find_folder(loc, date):
    '''
    Given the physical location and date of a form, get the appropriate archive
    destination path - if it doesn't exist yet, it will be created.
    
    Parameters:
        loc: physical location of the machine the form is on
        date: date of the form that is being moved
    Returns:
        new_path (str): file path that the form will be saved to
    '''
    
    #Get the date as a string
    date = date.strftime("%Y/%m/%d")
    #The new path in the archive
    new_path = os.path.join(archive_path, loc, date)
    #create folder for the date and intermediate folders if necessary
    os.makedirs(new_path, exist_ok = True)
    
    return new_path

#Logging format setup. Default level - WARNING, "w" - overwrite log
logging.basicConfig(filename=log_path, format="%(asctime)s - %(levelname)s: %(message)s",
                    datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO, filemode="w")

#Read the excel of machines
excel = pd.read_excel(machines_excel)
#locations = df.iloc[:, 0].tolist()
#serials = df.iloc[:, 6].tolist()
locations = ["Kilmarnock, VA", "Kinsale, VA", "Hurlock, MD", "Cofield, NC", "Cofield, NC", "Cofield, NC", "Cofield, NC"]
serials = ["2UA3110NCD", "2UA3110NCJ", "2UA55214W8", "2ua3110nbq", "2ua3110nby", "2ua3110nc2", "2ua3110ndx"]

#Go through the locations and serials together.
for (loc, serial) in zip(locations, serials):
    #If the machine pings
    if ping(serial):
        #Go to the forms folder and get all the files
        curr_path = "//4DLQ733/Mia/" + serial + "/c/Agris/datasets/001/Forms/"
        forms = os.listdir(curr_path)

        #Go through all the tickets.
        for form in forms:
            #Get ticket's date from its name.
            date = form.split('_')[6]
            date = datetime.strptime(date, '%y%m%d') #date as datetime object
            
            #If the form is older than the local max, archive it.
            if(datetime.today() - date > local_max and form not in archive):
                archive.append(form)

                #Find or create the archive destination based on the form's date
                archive_dest = find_folder(loc, date)
                #Copy to the archive
                shutil.copy(curr_path + form, archive_dest)
                #Track number archived
                num_archived += 1

                if log_details:
                    logging.info(form + " has been archived")

                #Remove form from local os.remove(local_path + "/" + form)
                #logging.info(form + " has been removed")
                #num_local_removed += 1

            else:
                logging.info("Form already in archive - " + form)

    else:
        #Log when a machine fails to ping
        logging.warning("Machine failed to ping - " + serial)

#Log script summary
logging.info("Total forms archived: " + str(num_archived))

