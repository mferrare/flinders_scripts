#
# Processes the Problem Statements CSV and creates channels for the team.
# Basically works as follows:
# - asks for the Team name
# - reads the CSV file
# - Iterates through CSV data.  Assumes the first column contains the
#   channel names (no sanity checking)
# - connects to teams and creates the channels in the specified team
#
# Note: Requires the MicrosoftTeam powershell module.  See here for
#       details: https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.4

# We assume that this will only ever be run from the current directory
# These are my general utilities
Import-Module .\MJFutils.psm1

#
# Main Program
#
function main() {
    $CSVpath, $CSVData = MJFGetCSVFile("Select CSV Export of Problem Statements")

    $theTeam = MJFGetTopicName

    # Connect to Teams
    $TeamCredentials = Connect-MicrosoftTeams

    # If the team doesn't exist we create it.
    $TopicTeam = Get-Team -DisplayName $theTeam
    if ( $TopicTeam -eq $null ) {
        $TopicTeam = New-Team -DisplayName $OverallTeamName
    }

    # Add the channels to the team
    # Note: This doesn't check to see whether the channel exists already nor
    # does it check for success/failure
    $CSVData | ForEach-Object {
        New-TeamChannel -GroupId $TopicTeam.GroupID -displayName $_.'Project Identifier'
    }
}

main
