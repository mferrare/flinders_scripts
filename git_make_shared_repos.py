# Reads through a FLO grades export and uses the export to create
# shared repositories and invite the group members as collaborators
from github import Github
import json, argparse, logging
from os import acces, R_OK
from os.path import isfile
import pandas as pd

def processCSV(pathToCSV):
    """
    Synopsis
    --------
    Processes the CSV and returns a list of hashes with relevant data

    Parameters
    ----------
    `pathToCSV` : path to CSV file.  It should be args.pathToFLOCSV but we pass it
    in as a parameter

    Description
    -----------
    The CSV file should contain at least two columns, being:
    - FAN
    - Group

    The Group column will be used to construct the repository name.  The FAN will be
    used to construct a list of list of GitHub users for each group

    Returns
    -------
    None or a (possibly empty) list of group/user data structures ie:
    - group name
        - member 1
        - member 2
        ...
        - member n
    - group name
        - member 1
        ...
    """

    try:
        csv_data = pd.read_csv(pathToCSV)
    except IOError as err:
        print "Error reading the file {}: {}".format(pathToCSV, err)
        return None
    
    # This is the result array
    result = {}

    for index, data in csv_data:
        # This serves as a sanity check on the columns in the CSV
        try:
            FAN = data.FAN
            MJFgroup = data.Group
        except AttributeError as err:
            print "CSV does not appear to have FAN and Group columns: {}".format(err)
            return None
        
        # If we're here we have correct column names.  Add a new entry to rou data structure
        # try the append first.  If it fails, then create the (new) record
        try:
            result[MJFgroup]
        except KeyError:
            result[MJFgroup] = []
        # Now we append
        result[MJFgroup].append(FAN)

    return result

 def createRepositories(topicCode, topicYear, topicSemester, groupData):
    """
    Synopsis
    --------
    Creates GitHub repositories and adds group members as collaborators

    Parameters
    ----------
    `topicCode` : str containing topic code (eg: 'LLAW3301')
    `topicYear` : str containing four digit year (eg: '2021')
    `topicSemester` : str containing two character topic semester (eg: 'S1')
    `groupData` : dict data structure from processCSV

    Description
    -----------
    - Uses the group name and the topic information to create a repository name
      Repo name example:  docassemble-LLAW33012020S1BDF1 (group name is BDF1)
     
    The CSV file should contain at least two columns, being:
    - FAN
    - Group

    The Group column will be used to construct the repository name.  The FAN will be
    used to construct a list of list of GitHub users for each group

    Returns
    -------
    None or a (possibly empty) list of group/user data structures ie:
    - group name
        - member 1
        - member 2
        ...
        - member n
    - group name
        - member 1
        ...
    """   

def main():
    """
    Main Program
    """
    global secrets
    global args

    # Process command-line arguments
    parser = argparse.ArgumentParser(description='Create shared repositories')
    parser.add_argument('--topicCode', required=True, help='Topic Code (eg: LLAW3301)')
    parser.add_argument('--topicYear', required=True, help='Topic year (eg: 2021)')
    parser.add_argument('--topicSemester', required=True, help='Topic Semester (eg: S1, NS1 etc)')
    parser.add_argument('--secrets_file', required=True, help='Path to secrets file (eg: /path/to/secrets.json')
    parser.add_argument('--secret', required=True, help='Name of secret to use')
    parser.add_argument('--pathToFLOCSV', require=True, help="Path to CSV file containing repo and group membership data")
    parser.add_argument('--loglevel', default=logging.INFO, choices=['INFO', 'WARN', 'DEBUG', 'ERROR'], help='Set logging level (default: INFO)')
    
    args = parser.parse_args()

    logging.basicConfig(level=args.loglevel)

    # Get the secrets
    secrets = MJFSecrets(args.secrets_file).getSecret(args.secret)
    logging.debug('secrets: {}'.format(secrets))

    # Process the CSV
    MJFGroupsList = processCSV(args.pathToFLOCSV)

    # Create the repo
