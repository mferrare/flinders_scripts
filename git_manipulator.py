from github import Github
import git
import os.path
import csv

# Global variables
g          = None    # PyGithub object
# Destination path for cloning assessments
clone_path = "C:/Users/ferr0182/Flinders/Flinders Law - LLAW3301 Law in a Digital Age/_LLAW3301_2020_S1/Assessment C3" 
# Single path - we're cloning all repos here and not into group dirs
single_path = "c:/users/ferr0182/OneDrive - Mark Ferraretto/Software/Flinders/GitHub"

# CSV Export of user data from FLO
user_data = "C:/Users/ferr0182/Flinders/Flinders Law - LLAW3301 Law in a Digital Age/_LLAW3301_2020_S1/llaw3301-law-in-a-digital-age---2020-s1 20200408.csv"

def initialise():
    # Expect: Nothing
    # Return: Nothing
    # Sets up our initial connections and populates global variables

    # Connect to my GH
    global g
    g = Github("91a44974b28eee9daa21e2cc9e2797a15deefd0d")

def fork_repo(FAN):
    # Expect: FAN
    # Return: Repository object of forked repository
    # Constructs the path to the repository on GH, being
    # <FAN>/docassemble-<FAN>

    path_to_repo = '{}/docassemble-{}'.format(FAN, FAN)
    repo_to_fork = g.get_repo(path_to_repo)
    return g.get_user().create_fork(repo_to_fork)

def clone_to_one_dir(forked_repo):
    
    # Clone into directory.
    clone_URL = forked_repo.html_url
    clone_dir = clone_URL.split('/')[-1]
    # We don't need to store the resut of the clone_from() call but
    # we do just in case we want it in future and also helps with
    # debugging.
    cloned_repo = git.Repo.clone_from(clone_URL, os.path.join(single_path, clone_dir))

def clone_repo(forked_repo, group):
    # Expect: Repository object containing clone information,
    #         Project group identifier
    # Return: Nothing
    # TODO: Return true/false on cloning?
    # Clones the repo into a subdir in the project dir

    # Make sure we have a group.
    if not group:
        raise "Group must have a value"

    project_dir = os.path.join(clone_path, group)
    # Make the directory
    try:
        os.mkdir(project_dir)
    except FileExistsError as e:
        # We don't care if the directory exits and is a directory
        if os.path.isdir(project_dir):
            pass
        else:
            raise e
    except Exception as e:
        raise e

    # Clone into directory.
    clone_URL = forked_repo.html_url
    clone_dir = clone_URL.split('/')[-1]
    # We don't need to store the resut of the clone_from() call but
    # we do just in case we want it in future and also helps with
    # debugging.
    cloned_repo = git.Repo.clone_from(clone_URL, os.path.join(project_dir, clone_dir))

    
def process_csv():
    # Expect: Nothing
    # Return: dictionary containing csv data we need
    # Opens CSV and returns a dictionary of FANs and project groups
    csvfile = open(user_data)
    reader = csv.DictReader(csvfile)

    result = {}
    for row in reader:
        FAN = row['FAN']
        result[FAN] = None
        # Search through the Groups until we find a group containing a
        # a project group.  We assume only project groups start with
        # 'P'
        for key in row.keys():
            try:
                if key[0:5] == 'Group' and row[key][0] == 'P':
                    result[FAN] = row[key]
                    break
            except IndexError as e:
                # Catch index out of range errors
                pass

    csvfile.close()
    return result

#
# Main Program
#
def main():
    initialise()

    # Get list of students and project groups
    all_data = process_csv()

    # For each student we fork their repo and clone to our path
    for FAN in all_data.keys():
        try:
            forked_repo = fork_repo(FAN)
            #clone_repo(forked_repo, all_data[FAN])
            #clone_to_one_dir(forked_repo)
        except Exception as e:
            print('Error processing FAN: {} group {}: {}'.format(FAN, FAN, str(e)))


if __name__ == "__main__":
    main()
