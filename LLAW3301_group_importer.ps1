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

    # Where we will store the CSV to import
    $outputData = @()


    # Get the CSV data
    $CSVpath, $CSVdata = MJFGetCSVFile -displayText "The file should be called 'W01 Seminar Allocations - group import to FLO.csv'"

    # Append to array formatted for import into FLO
    $CSVdata | ForEach-Object {
        $groupname = $_.'Project ID'
        $the_students = $_.'Members'
        $the_students -split "`n" |
            Select-String -CaseSensitive -Pattern '^.*\((?<fan>.+)\).*$' |
            ForEach-Object {
                $the_fan = $_.Matches[0].Groups['fan'].Value
                $a_element = New-Object PSObject -Property @{
                    groupname = $groupname
                    username  = $the_fan
                }
                $outputData += $a_element
            }
        }
    $outputData
    $outputData | Export-Csv -Path ./output.csv -NoTypeInformation
}

main