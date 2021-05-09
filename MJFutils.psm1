Add-Type -AssemblyName System.Windows.Forms

function MJFGetFilePath {
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
    [cmdletbinding()]
    param(
        [hashtable] $h_Properties
    )

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

function MJFGetTeamName() {
 <#
    .Synopsis
    Gets the Team name and returns it as a string

    .Description
    Dialog box to prompt user for team name.  This should
    be the full team name for the semester eg: LLAW3301_2020_S1.  The 
    team name is returned as a string.  If nothing is entered then '' (empty string)
    is returned
#>
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

function MJFMakeDir($s_dirPath, $s_newDir){
<#
    .Synopsis
    Makes a new directory in the specified parent

    .Parameter s_dirPath
    Path to parent in which new directory will be made

    .Parameter s_newDir
    Name of the new subdirectory to be created

    .Description
    Creates a new subdirectory in the specified parent path
#>

    # Make the path to PDFs
    $parentPath = Split-Path -Path $s_dirPath -Parent


    $PDFpath = Join-Path -Path $parentPath -ChildPath $s_newDir

    # Check if it's a directory first
    if ( Test-Path -Path $PDFpath -PathType Container ) {
        return $PDFpath
    }

    # If we're here it either doesn't exist or isn't a directory.  If it exists and
    # is a file then return $false.
    if ( Test-Path -Path $PDFpath -PathType Leaf ) {
        return $false
    }

    # If we're here then the path doesn't exist.  Create it
    $null = New-Item -Path $parentPath -Name $s_newDir -ItemType "directory"

    # Test again to make sure it's there
    # Check if it's a directory first
    if ( Test-Path -Path $PDFpath -PathType Container ) {
        # Returning $PDFpath is also true
        return $PDFpath
    } else {
        return $false
    }
}

function MJFGetDirPath ( $h_Properties )
{
<#
    .Synopsis
    Get the path of a selected directory

    .Description
    Makes a FileBrowser object, with the specified properties, for
    opening a file.  We select one file in the directory and the path
    of the directory containing the file is returned.  We select a directory
    this way because the DirectoryBrowser object is more difficult to use
    than a FileBrowser

    .Parameter $h_Properties
    A hash of properties
    - key is the name of the property
    - value is the value to set that property
#>
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Multiselect = $false
    }
    foreach($key in $h_Properties.keys) {
        $FileBrowser.$key = $h_Properties.$key
    }
    $null = $FileBrowser.ShowDialog()

    # We just want the directory name
    Split-Path -Path $FileBrowser.FileName
}

function MJFConvertToDate ( $dateString ) {
    <#
    .Synopsis
    Convert a date string received from a FLO marking CSV to a dateTime object and returns
    that object

    .Parameter dateString
    Date as string.  Should be formatted eg: Thursday, 12 September 2019, 4:00 PM
    #>
    # TODO: Sanity checking
    # We just assume the string is formatted as we want it
    [dateTime] $dateString.split(',',2)[1].trim()
}

Export-ModuleMember -Function MJFGetFilePath
Export-ModuleMember -Function MJFGetCSVFile
Export-ModuleMember -Function MJFGetTeamName
Export-ModuleMember -Function MJFMakeDir
Export-ModuleMember -Function MJFGetDirPath
Export-ModuleMember -Function MJFConvertToDate