# A script to notify users on LLAW3301 that their accounts
# will be deleted.  Uses the DA API to get a list of emails
# and smtplib and email to send the email.  Connects to my 
# SNS server
#
# I never finished this.

import json
import os
import requests

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

# List users
list_users_api = api_root + '/user_list'
# Let's do it!
payload = { 'key' : api_key }
r = requests.get(list_users_api, params=payload)
response = r.json()

for item in response:
    print("Email address: ", item['email'])


print("done")
