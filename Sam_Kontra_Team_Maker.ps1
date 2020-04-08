# Imports the project groups into FLO.  The CSV is formatted as follows
# - Column1: 'Project Identifier' - string
# - Column2: 'Project Description' - string
# - Column3: 'Allocated Seminar' - string
# - Column4: 'Members' - multiline string
#
# Column4 is the problem. Each user is listed separated by newlines and with
# their FANs in braces.  I can't go back to DA to get the seminars because
# it will re-shuffle them.  So, here I am!

Import-Module .\MJFutils.psm1

# Main Program
function main() {
    $CSVpath, $CSVData = MJFGetCSVFile("Select CSV Export of Topic Participants")

    $theTeam = MJFGetTeamName

    # Connect to Teams
    $TeamCredentials = Connect-MicrosoftTeams

    $TopicTeam = Get-Team -DisplayName $theTeam
    if ( $TopicTeam -eq $null ) {
        $TopicTeam = New-Team -DisplayName $OverallTeamName
    }

    # Add each member to the team
    $CSVData | ForEach-Object {
        $email = $_.'FAN' + '@flinders.edu.au'
        Add-TeamUser -GroupId $TopicTeam.GroupID -User $email
    }
}

main