# A script that allows an admin to push files to the current user's
# playground

import requests, json
import os
from os import listdir
from os.path import isfile, join
import argparse

# Global variable - store secrets and other config info in here.
config_data = {}

# Default project in the playground
project_name = "C3"
# List of folders to iterate through.  This list comes from the
# API documentation: https://docassemble.org/docs/api.html#playground_get
project_folders = [ 'questions', 'sources', 'static', 'templates', 'modules']

# Sends one file to the playground
def MJFpush_to_playground(MJFpayload):

    # Initialise    
    construct_API_payload(MJFpayload)

    # create the list of multiple files
    file_payload = []
    for file in MJFpayload['files']:
        file_payload.append(
            ('file', open(file, 'rb'))
        )
    # Send the file
    response = requests.post(MJFpayload['URL'], data=MJFpayload, files=file_payload)
    return response

# Lists files in the playground
def MJFlist_playground_files(MJFpayload):

    # Add API key etc to payload
    construct_API_payload(MJFpayload)

    # Make the request
    response = requests.get(MJFpayload['URL'], params=MJFpayload)

    # If the response code is 200 we're OK.  In which case return the 
    # list of files
    if response.ok:
        return response.json()
    else:
        return []

def MJFdelete_playground_file(MJFpayload):
    # Add API key, endpoint etc to payload
    construct_API_payload(MJFpayload)
    # Make the request
    response = requests.delete(MJFpayload['URL'], params=MJFpayload)
    if response.ok:
        return True
    else:
        return False
        


# List all the files in the specified directory
def MJFlist_files_in_dir(MJFdir):
    try: 
        return [join(MJFdir, f) for f in listdir(MJFdir) if isfile(join(MJFdir, f))]
    except:
        # Return nothing if the directory doesn't exist
        return []

# Get API key
# Expect: Nothing
# Return: Nothing
# API key is stored in a secrets file which is HARD coded in this
# sub.  Extract it and store it in config data global variable.
def get_API_key():
    # Secrets file
    userprofile = os.environ['USERPROFILE']
    secrets_file = userprofile + '/OneDrive - Mark Ferraretto/kdb/secrets.json'

    # Read the secrets in from the secrets file
    with open(secrets_file, 'r') as fp:
        secrets = json.loads(fp.read())
    fp.close()
    config_data['api_key'] = secrets['flinders_api_unrestricted']

# Get the API URL and store it in config_data
# Expect: Nothing
# Return: Nothing
# This was reading from a config file before but now it is
# HARD CODED in here.
# TODO: Should this get moved to a config file?
def get_API_URL():
    config_data['api_root'] = 'https://docassemble.flinders.edu.au/api'

# construct_API_payload
# Expect: API endpoint, dictionary with parameters to pass through to API
# Return: payload dictionary
# Constructs payload to pass to API.  
def construct_API_payload(payload):

    # Add URL and api key to the payload
    payload['URL'] = config_data['api_root'] + '/playground'
    payload['key'] = config_data['api_key']
    payload['project'] = project_name

# clean_out_playground
# Expect: Nothing
# Return: Nothing
# Deletes all files in the default project in the playground
def clean_out_playground():
    # Clean out the project area
    # Need to do this folder by folder.  We get a list of what's in each
    # folder and the delete it
    
    for folder in project_folders:
        payload = {}
        payload['folder'] = folder
        the_files = MJFlist_playground_files(payload)
        for a_file in the_files:
            payload = {}
            payload['folder'] = folder
            payload['filename'] = a_file
            MJFdelete_playground_file(payload)

# Main Program
def main():
    global project_name

    # Process command-line arguments
    parser = argparse.ArgumentParser(description='Push local code to playground')
    parser.add_argument('path_to_package', help='Path to the top of docassemble package (eg: /path/to/docassemble-packagename)')
    parser.add_argument('--project', '-p', help='Docassemble playground project name (default is {})'.format(project_name))

    args = parser.parse_args()

    # Store variables
    da_package = args.path_to_package
    if args.project is not None:
        project_name = args.project.strip()
    
    #Initialise
    get_API_key()
    get_API_URL()
    clean_out_playground()

    # We assume that the path supplied and now in da_package is of the format 
    # /path/to/docassemble-packagename
    # We need to construct the package name and we do some basic sanity checking
    # as well.
    # Get the last part of the path
    packagename = os.path.basename(os.path.normpath(da_package))
    # trim off the leading 'docassemble-'
    packagename = packagename.replace('docassemble-', '')

    # Construct the path to the four folders
    path_to_folders = os.path.join(da_package, 'docassemble/{}/data'.format(packagename))

    # Now we can push files up.  We only push up stated folders
    for a_folder in project_folders:
        MJFpayload = {}
        MJFpayload['folder'] = a_folder

        # Construct the list of files and add to payload
        folder_path = os.path.join(path_to_folders, a_folder)
        MJFpayload['files'] = MJFlist_files_in_dir(folder_path)

        # Send
        MJFresponse = MJFpush_to_playground(MJFpayload)
        if MJFresponse.ok:
            print('Pushed {}'.format(path_to_folders))
        else:
            print("Error: pushing {} to package {}: {} {} {} ".format(a_folder, packagename, MJFresponse.status_code, MJFresponse.reason, MJFresponse.text))
main()