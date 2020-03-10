# A script that will allow for batch-adding users.
# Useful documentation is here: https://requests.readthedocs.io/en/master/user/quickstart/#make-a-request
# I never finished this.

import requests, json
import csv

def MJFmodify_privs(MJFpayload):
    # Create a user
    URL = api_root + '/user/new'
    # Construct the payload
    payload = {
        'key'        : api_key,
        'username'   : MJFpayload['email'],
        'privileges' : 'developer',
        'first_name' : MJFpayload['first_name'],
        'last_name'  : MJFpayload['last_name'],
        'timezone'   : 'Australia/Adelaide'
    }
    response = requests.post(URL, data=payload)
    return response

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
api_key  = secrets_j['llaw3301_api_key']

# email server
MJFemailserver = smtplib.SMTP('smtp.flinders.edu.au', 25)

# Get the CSV records
csv_file = userprofile + '/Flinders/Flinders Law - LLAW3301 Law in a Digital Age/_LLAW3301_2020_S1/Test - 1 record.csv'
with open(csv_file, 'r') as fp_csv:
    MJFrecords = csv.reader(fp_csv)
    for MJFrecord in MJFrecords:
        # Process each record

        # Skip the first line
        if MJFrecord[0] == 'FAN':
            continue

        # Now we process
        MJFitem = {}
        MJFitem['FAN'] = MJFrecord[0]
        MJFitem['first_name'] = MJFrecord[1]
        MJFitem['last_name'] = MJFrecord[2]
        MJFitem['email'] = MJFrecord[3]

        # Create the account
        response = MJFcreate_account(MJFitem)
        if response.status_code == 200:
            # We're OK.  Add the password to the record
            temp = response.json()
            MJFitem['password'] = temp['password']
            MJFitem['created'] = True
        else:
            MJFitem['created'] = False
            MJFitem['status'] = response.text
        
        # Now send an email.
        msg = str(MJFitem)
        result = MJFemailserver.sendmail('do-not-reply@flinders.edu.au', 'ferr0182@flinders.edu.au', msg)




print("Just a breakpoint")