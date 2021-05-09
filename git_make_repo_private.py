from github import Github

# Global variables
g          = None    # PyGithub object

def initialise():
    # Expect: Nothing
    # Return: Nothing
    # Sets up our initial connections and populates global variables

    # Connect to my GH
    global g
    g = Github("28329cd6600a3ae76b21bf5093669c26fcf64cf2")

initialise()
for repo in g.get_user().get_repos():
    if repo.name == 'docassemble-LLAW33012020S2P01':
        print(repo.name)
        print(repo.private)
        repo.edit(private=True)
        print(repo.private)