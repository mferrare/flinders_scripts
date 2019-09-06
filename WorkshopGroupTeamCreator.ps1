# Adds students to the main team and to each workshop group team as
# per the imported CSV file

$OverallTeamName = "LLAW2221_2019_S2"

$CSVData = Import-CSV 'C:\users\mark\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Group Allocations.csv'

# Connect to teams
$TeamCredentials = Connect-MicrosoftTeams

$TopicTeam = Get-Team -DisplayName $OverallTeamName
if ( $TopicTeam -eq $null ) {
    $TopicTeam = New-Team -DisplayName $OverallTeamName
}

# Run through the CSV file.  For each entry:
# - create the team if it doesn't exist
# - add the team member to the team
# We pass the list in sorted by group name.  As the group changes
# we create a new team
$oldgroup = ""
$groupid = ''
$CSVData | Sort-Object -Property group | ForEach-Object {
    $group = $_.group
    $fan = $_.username
    $email = $fan + "@flinders.edu.au"
    $GroupTeamName = $OverallTeamName + '_' + $group

    # If this group is different from the last one, create
    # a new team
    if ( $group -ne $oldgroup ) {
        $GroupTeam = New-Team -DisplayName $GroupTeamName
        Write-Host "Created Team: " $GroupTeam.DisplayName " id: " $GroupTeam.GroupID
    } else {
        $GroupTeam = $OldGroupTeam
    }
    Add-TeamUser -GroupId $GroupTeam.GroupID -User $email
    Add-TeamUser -GroupID $TopicTeam.GroupId -User $email
    Write-Host "Team: " $GroupTeamName " added: " $email
    $oldgroup = $group
    $OldGroupTeam = $GroupTeam
}
