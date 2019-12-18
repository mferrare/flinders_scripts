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

# MJFGetFilePath
# Expect: An hash of properties.
# - key is the name of the property
# - value is the value to set that property
# Return: Full path name to the file selected
# Makes a FileBrowser object with the specified properties for opening
# a file.  We want all filebrowsers to select one file only.
function MJFGetDirPath ( $h_Properties )
{
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

# MJFGetSubmissionDeadlines
# Expect: Nothing
# Return: a hash of submission deadline information
function MJFGetSubmissionDeadlines() {

    # Find the CSV file with the deadlines...
    $fp_Deadlines = MJFGetFilePath -h_Properties @{ 
        Title = "Find Deadlines File";
        InitialDirectory = "C:\Users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Assignments";
        Filter = 'CSV File (*.csv)|*.csv';
    }
    #$fp_Deadlines = "C:\Users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Assignments\Submission Deadline.csv"

    # ...and push the CSV into a hash and return it
    $h_Deadlines = @{ }
    Import-Csv -Path $fp_Deadlines | ForEach-Object { 
        $h_Deadlines[$_.Seminar] = @{ Day = $_.'Submission Deadline Day'; Time = $_.'Submission Deadline Time'; WordCount = $_.'Word Limit'  }
    }
    return $h_Deadlines
}

# MJFGetCSVFile
# Expect: Nothing
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

# MJFGetSubmissionFilePaths
# Iterate through the folder containing submissions.  Create a hash with
# key   : FAN
# value : Full path to word document
function MJFGetSubmissionFilePaths() {

    # return result in this hash
    $result = @{}

    # Get the folder where the documents are stored
    $WordDocFolder = MJFGetDirPath -h_Properties @{ Title = "Select folder containing assignment submissions" }
    #$WordDocFolder = "C:\users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Assignments\Workshop 2\submissions"

    # Iterate through the files in that folder
    Get-ChildItem $WordDocFolder -Filter *.docx | ForEach-Object {
        # Each file is named with this format
        # <FAN>_submission_<workshop group>-<full name>_assignsubmission_file_<document title>.docx
        # We look for valid files by checking the name conforms to the FAN_submission prefix. If so, we add it to
        # our hash. Otherwise skip.
        # A FAN is always four characters and four digits
        if ( $_.Name -match '[a-z][a-z][a-z][a-z][0-9][0-9][0-9][0-9]_submission' ) {
            $FAN, $nothing = $_.Name.Split('_',2)
            $result[$FAN] = $WordDocFolder + "\" + $_.Name
        }
    }
    return $result
}

# MJFWordCountPenalty
# Expect: Penalty percentage expressed as a decimal (eg: 5% == 0.05),
#         number of words as increment to step penalty (eg: 5% per 100 words)
#         word limit, maximum attainable marks
#         and (Word) file to apply it to.
# Return: An array being [ (float) penalty, (string) feedback]
# the feedback
# 1. Nominate the folder in which the submissions are - do this at the beginning?
# 2. how do we work out who owns what word file?  I think the FAN is used in the 
#    filename
# 3. Open and count as per https://devblogs.microsoft.com/scripting/weekend-scripter-use-powershell-to-count-words-and-display-progress-bar/
function MJFWordCountPenalty ( $penaltyDecimal, $penaltyWordIncrement, $wordLimit, $maximumMarks, $submissionFilePath ) {

    # Get the word count from the document
    if ( $submissionFilePath ) {
        #TODO: exception handling when opening non-existent or invalid documents
        try {
            $wordDoc = $MSWord.Documents.Open($submissionFilePath)
            $wordCount = $MSWord.ActiveDocument.ComputeStatistics(0, $false)
        }
        catch {
            Write-Host "Error opening document: $submissionFilePath"
        }
        finally {
            $wordDoc.close() | Out-Null
        }
    }

    # TODO: What if word limit makes no sense?

    # Feedback always starts with this
    $lateFeedback = ""

    # Calculate the penalty
    if ( $wordCount -le $wordLimit ) {
        $wordPenalty = 0
    } else {
        # Calculate the penalty as the number of words over the limit divided by the increment
        $wordsOverLimit = $wordCount - $wordLimit
        $wordsOverIncrement = [math]::Round( ($wordsOverLimit / $penaltyWordIncrement) + 0.5)
        $wordPenalty = $wordsOverIncrement * $maximumMarks * $penaltyDecimal
        $lateFeedback = "Word count: $wordCount  Limit: $wordLimit"
    }

    # Here's what to return
    $wordPenalty
    $lateFeedback
}

#MJFLateSubmissionPenalty
# Expect: submission deadline, submission DateTime, max marks attainable, penalty to apply to max marks (percentage expressed as a decimal)
# Return: array [ (float) penalty, (string) feedback ]
# Calculates lateness of submission and applies penalty if so.
function MJFLateSubmissionPenalty ( $SubmissionDeadline, $SubmissionDateTime, $maximumMarks, $penaltyDecimal ) {

    # Feedback to return
    $lateFeedback = ""

    # Calculate submission lateness
    $d_submissionLateness = $SubmissionDeadline - $SubmissionDateTime
    if ( $d_submissionLateness.TotalHours -lt 0 ) {
        # We reduce the mark by the specified percentage per day
        $f_reduction = [math]::Round(($d_submissionLateness.TotalHours / 24 * -1 ) + 0.5) * $maximumMarks * $penaltyDecimal
        $lateFeedback += "Submission deadline: $SubmissionDeadline Submitted: $SubmissionDateTime"
    } else {
        $f_reduction = 0.0
    }

    # Return
    $f_reduction
    $lateFeedback
}

# MJFConvertToDate
# Expect: String which is field from marking CSV.  Should be formatted eg: Thursday, 12 September 2019, 4:00 PM
# Return: Date as DateTime
function MJFConvertToDate ( $dateString ) {
    # TODO: Sanity checking
    # We just assume the string is formatted as we want it
    [dateTime] $dateString.split(',',2)[1].trim()
}

# MJFGetMarksHash
# Expect: Nothing
# Return: hash table formatted as below
# Note:  I have to change how this hash is created each time as the marking criteria will change
function MJFGetMarksHash ( $displayText ) {

    $hashTable = @{}
    $s_filePath, $a_table = MJFGetCSVFile -displayText $displayText
    $a_table | ForEach-Object {
        $hashTable[$_.FAN] = $_
    }

    # I don't think we need the file path.  Just return the hash
    $hashTable
}

# MJFWriteFeedback
# Expect: marking record hash, late penalty string, word count penalty string, late penalty, word count penalty
# Return: Feedback string to insert into grades CSV
# This function needs to be customised to reflect each individual assessment
# as assessment criteria (and therefore object names) will vary from
# assessment to assessment.
function MJFWriteFeedback ( $currentRecord, $latePenalty, $latePenaltyString, $wcPenalty, $wcPenaltyString ) {

    $c1 = $currentRecord.Research_9
    $c2 = $currentRecord.Analysis_13
    $c3 = $currentRecord.Structure_3
    $c4 = $currentRecord.AGLC_3
    $c5 = $currentRecord.Readability_2
    $grade = $currentRecord.Grade
    $feedback = $currentRecord.'Feedback Comments'

    $allFeedback = "<p><b>Criterion 1 (Research)</b>                  : $c1</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 2 (Analysis)</b>   : $c2</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 3 (Structure)</b>  : $c3</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 4 (AGLC)</b>       : $c4</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 5 (Readability)</b>: $c5</p>"

    if ( $latePenalty + $wcPenalty -gt 0 ) {
        $netgrade = $grade - $latePenalty - $wcPenalty
        $allFeedback = $allFeedback + "<p></p><p><b>Mark awarded (before penalties)</b> : $grade</p>"
        if ( $latePenalty -gt 0 ) {
            $allFeedback = $allFeedback + "<p><b>Late penalty</b>      : $latePenalty ($latePenaltyString)</p>"
        }
        if ( $wcPenalty -gt 0 ) {
            $allFeedback = $allFeedback + "<p><b>Word count penalty</b>: $wcPenalty ($wcPenaltyString)</p>"
        }
    } else {
        $netgrade = $grade
    }
    $netpercent = "{0:P0}" -f ($netgrade / $currentRecord.'Maximum Grade')
    $allFeedback = $allFeedback + "<p><b>Mark awarded:</b>           : $netgrade</p>"
    $allFeedback = $allFeedback + "<p><b>Mark as percentage:</b>     : $netpercent</p>"

    if ( $feedback -ne "" ) {
        $allFeedback = $allFeedback + "<p>$feedback</p>"
    }
    # Return the feedback
    $allFeedback
}

#
# Main Program
# 
function main() {
    
    # Get our data
    $s_gradesPath, $a_gradesTable = MJFGetCSVFile -displayText "Select the Grades as exported from FLO"
    $h_MarksHash = MJFGetMarksHash -displayText "Select the CSV where you have entered your marks"
    $h_SubmissionFileNames = MJFGetSubmissionFilePaths

    # We need to open up an instance of Word for word counting.
    $MSWord = New-Object -ComObject word.Application
    $MSWord.visible = $false

    # Iterate through marking table and process each row as we go
    $i_index = 1    # We just use this to display a basic progress indicator
    foreach ( $row in $a_gradesTable) {
        # Store FAN for easy reference
        $s_currentFAN = $row.FAN

        # Write some output to the console
        Write-Host "Processing $s_currentFAN ( $i_index /"$a_gradesTable.count")"

        $i_index++

        # If there's nothing in the marking hash then we can skip this entry
        if ( -Not $h_MarksHash.ContainsKey($s_currentFAN) ) {
            Write-Host "$s_currentFAN - no entry in marking spreadsheet"
            continue
        }

        # Skip 'No submission' entries
        if (  $row.'Individual submission status' -match '^No submission.*' ) {
            Write-Host "$s_currentFAN - No submission"
            continue
        }
    
        # Late penalty calculation
        # Convert the date in the marking CSV to a dateTime
        $d_currentSubmission = MJFConvertToDate -dateString $row.'last modified (submission)'
        # ...and get the deadline
        $d_submissionDeadline = MJFConvertToDate -dateString $row.'Due Date'

        # Get late penalty
        # TODO: This doesn't work for group submissions as the 'Due Date' field does not
        # exist in the CSV file!
        $latePenalty, $lateFeedback = MJFLateSubmissionPenalty -SubmissionDeadline $d_submissionDeadline `
                                        -SubmissionDateTime $d_currentSubmission `
                                        -maximumMarks $row.'Maximum Grade' `
                                        -penaltyDecimal 0.05 
        $latePenalty = 0
        $lateFeedback = ""
        
        # Get Word count penalty
        $wordPenalty, $wcFeedback = MJFWordCountPenalty -penaltyDecimal 0.05 `
                                        -wordLimit 1500 `
                                        -submissionFilePath $h_SubmissionFileNames[$s_currentFAN] `
                                        -maximumMarks $row.'Maximum Grade' `
                                        -penaltyWordIncrement 100

        # Apply penalties to grade
        $row.Grade = $h_MarksHash[$s_currentFAN].Grade
        $row.Grade -= $latePenalty
        $row.Grade -= $wordPenalty
        # Can't have a score less than zero.
        if ( $row.Grade -lt 0 ) {
            $row.Grade = 0
        }

        # Update and write feedback
        $row.'Feedback comments' = MJFWriteFeedback -currentRecord $h_MarksHash[$s_currentFAN] `
                            -latePenalty $latePenalty `
                            -latePenaltyString $lateFeedback `
                            -wcPenalty $wordPenalty `
                            -wcPenaltyString $wcFeedback
    }

    # Write the CSV
    $outputFile = $s_gradesPath + "_new.csv"
    $a_gradesTable | Export-CSV -Path $outputFile -NoTypeInformation

    # Clean up
    $MSWord.Quit()
    Remove-Variable MSWord
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}

main