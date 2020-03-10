Add-Type -AssemblyName System.Windows.Forms

function MJFGetFilePath ( $h_Properties )
{
    <#
        .Synopsis
        Gets full path of the selected file
    
        .Description
        Presents a dialog to select a file.  Properties can be passed in to the
        function to manage how the file dialog is displayed.  Only a single file can
        be selected

        .Parameter h_Properties
        A hash of properties.  key is the name of the property and value is the what the property should be set to
    #>
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Multiselect = $false
    }
    foreach($key in $h_Properties.keys) {
        $FileBrowser.$key = $h_Properties.$key
    }
    $null = $FileBrowser.ShowDialog()

    return $FileBrowser.FileName
}


function MJFGetCSVFile {
<#
    .Synopsis
    Open CSV file and return its path and data

    .Description
    Presents a FileBrowser dialog to select a CSV file.  Opens the file
    and reads it into a variable.  The path to the CSV and the CSV data
    are returned

    .Parameter displayText
    Text to display in FileBrowser dialog
#>
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)][String]$displayText
    )

    # Code goes here!
    $s_filePath = MJFGetFilePath -h_Properties @{
        Title = $displayText
        Filter = 'CSV File (*.csv)|*.csv'
    }

    $t_asTable = Import-Csv -Path $s_filePath

    $s_filePath
    $t_asTable
}

<#
    .Synopsis
    Gets the Team name and returns it as a string

    .Description
    Dialog box to prompt user for team name.  This should
    be the full team name for the semester eg: LLAW3301_2020_S1.  The 
    team name is returned as a string.  If nothing is entered then '' (empty string)
    is returned
#>
function MJFGetTeamName() {
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

Export-ModuleMember -Function MJFGetCSVFile
Export-ModuleMember -Function MJFGetTeamName