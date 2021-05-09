from github import Github

g = Github("28329cd6600a3ae76b21bf5093669c26fcf64cf2")

repos = { 
    "BDF1" : [],
    "FLAC2" : [],
    "CBS1"  : [],
    "FLPN2" : [],
    "FLAC1" : ['mitc0398', 'gold0230', 'lync0120', 'egar0007'],
    "FLN2"  : [ 'prie0075', 'symo0097', 'prat0105', 'macc0016', 'miln0124'],
    "HSC1"  : ['gues0024', 'neum0031', 'lave0072', 'spen0171'],
    "RSB1"  : ['groo0048', 'royc0004', 'webs0093', 'khou0029', 'gibb0151'],
    "RSPCA1": ['coll0020', 'habi0041', 'brad0318', 'cave0023', 'last0012', 'mule0013', 'chun0153'],
    "TIAS1" : ['manu0083', 'poli0048', 'tanu0011', 'carl0118', 'hugh0269', 'thom1338'],
    "VSS1"  : ['gold0255', 'mand0117', 'broo0407', 'geor0287'],
    "WWC1"  : ['lau0252', 'pott0198', 'fors0097', 'luu0016', 'foto0006']
}

for repo in repos:
    print("Group: {}, repo: https://github.com/LLAW3301/docassemble-LLAW33012021S1{}".format(repo, repo))


# for repo in repos:
#     MJFrepo = g.get_repo('LLAW3301/docassemble-LLAW33012021S1{}'.format(repo))
#     for member in repos[repo]:
#         try:
#             MJFrepo.add_to_collaborators(member, permission='push')
#         except Exception as e:
#             print("Unable to add FAN: {} to group repo: {} Error: {}".format(member, repo, e))

