$FilePath = "C:\Users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Group Allocations summary.csv"

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