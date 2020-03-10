# A script that allows an admin to push files to the current user's
# playground

import requests, json
import os
from os import listdir
from os.path import isfile, join

# Sends one file to the playground
def MJFpush_to_playground(MJFpayload):
    # Construct the payload
    payload = {
        'key'        : MJFpayload['api_key'],
        'folder'     : MJFpayload['folder'],
        'project'    : MJFpayload['project']
    }
    # create the list of multiple files
    file_payload = []
    for file in MJFpayload['files']:
        file_payload.append(
            ('file', open(file, 'rb'))
        )
    # Send the file
    response = requests.post(MJFpayload['URL'], data=payload, files=file_payload)
    return response

# List all the files in the specified directory
def MJFlist_files_in_dir(MJFdir):
    return [join(MJFdir, f) for f in listdir(MJFdir) if isfile(join(MJFdir, f))]


# Main Program
def main():

    # Secrets file
    userprofile = os.environ['USERPROFILE']
    secrets_file = userprofile + '/OneDrive - Mark Ferraretto/kdb/secrets.json'

    # Read the secrets in from the secrets file
    with open(secrets_file, 'r') as fp:
        secrets = fp.read()
    fp.close()
    secrets_j = json.loads(secrets)

    # Read in config file
    with open('docassemble_notify_users.json', 'r') as fp:
        config_text = fp.read()
    config = json.loads(config_text)

    # Store config variables somewhere easier to read
    api_root = config['api_root']
    api_key  = secrets_j['flinders_api_key']

    # Construct the payload
    MJFpayload = {
        'URL'       : api_root + '/playground',
        'api_key'   : api_key,
        'project'   : 'sysadmin',
        'folder'    : 'questions',
    }
    # Construct the list of files
    # only add .yml files
    all_files = MJFlist_files_in_dir(MJFdir)
    push_files = []
    for a_file in all_files:
        if a_file[-4:] == '.yml':
            push_files.append(a_file)
    MJFpayload['files'] = push_files

    # Send
    MJFresponse = MJFpush_to_playground(MJFpayload)
    if MJFresponse.status_code == 204:
        print("All good!")
    else:
        print("Error: ", MJFresponse.text)

MJFdir = 'C:/Users/ferr0182/OneDrive - Mark Ferraretto/Software/Flinders/GitHub/docassemble-sysadmin/docassemble/sysadmin/data/questions'
#MJFdir = 'C:/Users/Mark/Onedrive - Mark Ferraretto/Software/Flinders/GitHub/flinders_scripts'
main()