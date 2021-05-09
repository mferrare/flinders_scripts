#
# Processes the group import file and creates the private channels and
# adds members to each private channel.
# Basically works as follows:
# - asks for the Team name
# - reads the CSV file
#   - CSV file has two columns 'username' and 'group'
#   - columns must be in that order.  THERE IS NO SANITY CHECKING
# - The group name is the channel name.
# - The username is the FAN
#
# Note: Requires the MicrosoftTeam powershell module.  See here for
#       details: https://www.powershellgallery.com/packages/MicrosoftTeams

# We assume that this will only ever be run from the current directory
# These are my general utilities
Import-Module .\MJFutils.psm1

#
# Main Program
#
function main() {
    $CSVpath, $CSVData = MJFGetCSVFile("Select FLO Group Import CSV")

    $theTeam = MJFGetTeamName

    # Connect to Teams
    $TeamCredentials = Connect-MicrosoftTeams

    # If the team doesn't exist we create it.
    $TopicTeam = Get-Team -DisplayName $theTeam
    if ( $null -eq $TopicTeam ) {
        $TopicTeam = New-Team -DisplayName $OverallTeamName
    }

    # Add the channels to the team
    # Note: This doesn't check to see whether the channel exists already nor
    # does it check for success/failure
    $CSVData | ForEach-Object {
        # See if the team exists first
        $theChannels = Get-TeamChannel -GroupId $TopicTeam.GroupID
        # Add the channel if it doesn't exist
        if ( -not $theChannels.DisplayName -contains $_.'group' ) {
            $teamId = New-TeamChannel -GroupId $TopicTeam.GroupID -displayName $_.'group' -MembershipType Private
        } else {
            $position = $theChannels.DisplayName::indexof($theChannels.DisplayName, $_.'group')
            $teamId = $theChannels[$position].Id
        }
        # And add the member
        $newMember = -join($_.'username', "@flinders.edu.au")
        Add-TeamChannelUser -GroupId $TopicTeam.GroupID -DisplayName $_.'group' -user $newMember

    }
}

main
