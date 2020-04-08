#
# When downloading group assignment submissions from FLO, multiple copies of the 
# same submission are downloaded, one for each group member.  This makes marking
# and uploading feedback tedious and problematic.
#
# This script works through all the submitted files and prepares a subset of files
# to mark, which are copied into a separate directory.  The FANs for each group are
# also recorded so that, after marking, a copy of feedback files can be uploaded
# to each group member.

Import-Module .\MJFutils.psm1

# Initialise variables
$hashOfSubmissions = @{}

function constructHashOfSubmissions {
<#
    .Synopsis
    Iterate over all the word files in the submissions directory
    and construct a hash containing all the information we need
    to construct the consolidated list.
#>
    # Get the path to the dir where the files to process are
    $submissionDir = MJFGetDirPath -h_Properties @{ 
        Title = "Select any Word document in the Submissions folder";
        Filter = 'Word Files (*.doc?)|*.doc?';
    }

    # Work through each file in that directory
    $itemList = Get-ChildItem -Path $submissionDir -Filter *.doc? 
    $itemList | ForEach-Object {
        $_.Name | Select-String -Pattern '^([^_]+)_[^_]+_([^-]+)-(.*).docx$' | 
            ForEach-Object { 
                # Set temp variables
                $fileName, $FAN, $groupName, $leftOver = $_.Matches[0].Groups[0..3].Value 

                # Check if this file is a biblio
                $isBiblio = $leftOver -match 'biblio'

                # Initialise the hash if it doesn't exist
                if ( ! $hashOfSubmissions.ContainsKey($groupName) ) {
                    # Initialise this
                    $hashOfSubmissions[$groupName] = New-Object PSObject -Property @{
                        biblioFile = $false
                        submissionFile = $false
                        FANs = @()
                    }
                }

                # We're here so the hash exists.  Populate it.

                # If this is a submission file path and we don't already have a file populated
                # then add this file path as the submission file
                if ( !$isBiblio -and !$hashOfSubmissions[$groupName].submissionFile ) {
                    $hashOfSubmissions[$groupName].submissionFile = $filename
                }
                if ( $isBiblio -and !$hashOfSubmissions[$groupName].biblioFile ) {
                    $hashOfSubmissions[$groupName].bibliofile = $fileName
                }
                $hashOfSubmissions[$groupName].FANs += $FAN
            }
    }
}

function main() {
<#
    .Synopsis
    Main program
#>
    constructHashOfSubmissions
    $hashOfSubmissions
}

# Call the main program
main