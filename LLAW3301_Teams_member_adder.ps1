Add-Type -AssemblyName System.Windows.Forms

# Global Variables

# MJFGetFilePath
# Expect: An hash of properties.
# - key is the name of the property
# - value is the value to set that property
# Return: Full path name to the file selected
# Makes a FileBrowser object with the specified properties for opening
# a file.  We want all filebrowsers to select one file only.
function MJFGetFilePath ( $h_Properties )
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Multiselect = $false
    }
    foreach($key in $h_Properties.keys) {
        $FileBrowser.$key = $h_Properties.$key
    }
    $null = $FileBrowser.ShowDialog()

    return $FileBrowser.FileName
}

# MJFGetCSVFile
# Expect: text to display in file selection window
# Return: path to CSV, CSV data imported into table
# Opens CSV file and returns its path and data in a table
function MJFGetCSVFile( $displayText ) {
    $s_filePath = MJFGetFilePath -h_Properties @{
        Title = $displayText
        Filter = 'CSV File (*.csv)|*.csv'
    }

    $t_asTable = Import-Csv -Path $s_filePath

    $s_filePath
    $t_asTable
}

# MJFGetTopicName
# A simple(!) input form to get the topic name
function MJFGetTopicName() {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Team Name'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Enter Team Name (eg: LLAW3301_2020_S1)'
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox)

    $form.Topmost = $true

    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $x = $textBox.Text
    } else {
        $x = ""
    }
    $x
}

#
# Main Program
#
function main() {
    $CSVpath, $CSVData = MJFGetCSVFile("Select CSV Export of Topic Participants")

    $theTeam = MJFGetTopicName

    # Connect to Teams
    $TeamCredentials = Connect-MicrosoftTeams

    $TopicTeam = Get-Team -DisplayName $theTeam
    if ( $null -eq $TopicTeam ) {
        $TopicTeam = New-Team -DisplayName $theTeam
    }

    # Add each member to the team
    $CSVData | ForEach-Object {
        $email = $_.'FAN' + "@flinders.edu.au"
        Write-Host "Adding:" $email
        Add-TeamUser -GroupId $TopicTeam.GroupID -User $email
    }
}

main
