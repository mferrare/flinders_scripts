from github import Github
import csv

g = None

# CSV Export of user data from FLO
user_data = "C:/Users/mark/Flinders/Flinders Law - LLAW3301 Law in a Digital Age/_LLAW3301_2020_S1/llaw3301-law-in-a-digital-age---2020-s1 20200408.csv"

# Root name of repository
repo_root = 'docassemble-LLAW33012020S1'

def initialise():
    # Expect: Nothing
    # Return: Nothing
    # Sets up our initial connections and populates global variables

    # Connect to my GH
    global g
    g = Github("xxx")

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
                    result[FAN] = row[key][0:3]
                    break
            except IndexError:
                # Catch index out of range errors
                pass

    csvfile.close()
    return result


# Main Program
#
def main():
    initialise()

    # Get list of students and project groups
    all_data = process_csv()

    user = g.get_user()

    # For each student we fork their repo and clone to our path
    for FAN in all_data.keys():
        # Only add to repos if we have a repo
        if all_data[FAN] is None:
            continue

        repo_name = repo_root + all_data[FAN]
        try:
            repo = user.get_repo(repo_name)
            # Add the user as a push, pull collaborator
            repo.add_to_collaborators(FAN, 'push')
            repo.add_to_collaborators(FAN, 'pull')
        except Exception as e:
            print("Unable to add" , FAN, "to repo: ", repo_name, "Error: ", str(e))

if __name__ == "__main__":
    main()
