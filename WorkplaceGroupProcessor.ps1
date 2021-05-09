# Creates a HTML table that displays workshop group allocations
#
# Takes the group allocations summary and prepares a HTML table which can
# be copied and pasted into a FLO label.
# 
# Note: $FilePath is hard-coded
# Note: this is one of the earliest scripts I wrote.  It's not very refined.  Be careful!
$FilePath = "C:\Users\markf\Flinders\Flinders Law - Documents\LLAW3312 Constitutional Law\_LLAW3312_2021_S1\LLAW3312_2021_S1 Workshop Group Allocations Table.csv"
$CSVData = Import-Csv -Path $FilePath

# This is a dictionary
$OutputData = @{}
# This is a list
$OutputList = @()

$CSVData | ForEach-Object {
    $groupname = $_.Subgroup
    $NameFAN = $_.NameFAN
    $OutputData.$groupname = $OutputData.$groupname + ", " + $NameFAN
}

# Table Header
Write-Host "<table border=1 width=60%>"

# Trim off the leading comma and print out the table
$OutputData.GetEnumerator() | Sort-Object -Property Name | ForEach-Object {
    $group = $_.Name
    $temp = $_.Value
    $temp = $temp.TrimStart(", ")
    Write-Host "<tr><td>" $group "</td><td>" $temp "</td></tr>"
}

# Table footer
Write-Host "</table>"