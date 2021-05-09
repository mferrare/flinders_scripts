# - Reads grade from a FLO CSV file.
# - Looks for submission word file and counts words and
#   adjusts penalties accordingly.
# - Reads attendance from a FLO attendance file
#   - attendance file is exported from FLO attendance widget as xlsx
#   - top three rows of spreadsheet are removed (so only the table remains)
#   - file is saved as CSV
#
# NOTE: 
#  THIS VERSION OF MARKONATOR DOES NOT CALCULATE LATE PENALTIES
#  The functions are still in here but they are not used as A1 assessments
#  cannot be submitted late.
#

Import-Module .\MJFutils.psm1

# Global Variables
$overword_limit = 500
$overword_increment = 100
$overword_penalty = 0.05



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

function MJFGetAttendanceData() {
<#
    .Synopsis Get attendance data from spreadsheet exported from FLO Attendance Widget

    .Description
    THIS IS A FRAGILE FUNCTION.  It is not smart about processing the excel spreadsheet.
    It assumes that the FAN will exist in $col_fan (an integer) and the attendance data
    will exist in $col_att

    BOTH THESE COLUMN INDICES SHOULD BE CHECKED BEFORE RUNNING THIS FUNCTION

    This function also assumes that there will only ever be one worksheet in the
    Excel workbook and that the table ends at the first empty row.

    All that aside, it's pretty robust! :-(

    Returns a hash of FANs and attendance (as a boolean)
#>
    # Column indices for FAN and attendance
    $col_fan = 2
    $col_att = 14
    # First row of table
    $first_row = 4

    $fp_Attendance = MJFGetFilePath -h_Properties @{
        Title = "Where is the Attendance Spreadsheet?";
        Filter = 'SpreadSheet (*.xlsx)|*.xlsx';
    }
    #$fp_Attendance = "C:\Users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Assignments\Workshop 2\Workshop 2 Attendance - Test.xlsx"

    # Get the Attendance records.  This is more difficult as it's Excel
    $excel = New-Object -Com Excel.Application
    $wb_Deadlines = $excel.Workbooks.Open($fp_Attendance)
    # We assume there's only one sheet...
    $sh_Deadlines = $wb_Deadlines.Sheets.Item(1)
    # ...and the table starts in cell A4
    $cel_Start = $sh_Deadlines.Cells.Item($first_row, 1)
    # TODO: Check the cell value contains Student ID.  Otherwise we're reading the wrong spreadsheet
    $null = $cel_Start.Value2

    # Read through the data in the table
    # We assume that the first empty row signifies the end of the table
    $h_AttendanceInfo = @{ }
    $i = 5
    while ($true) {
        $cel_FAN = $sh_Deadlines.Cells.Item($i, $col_fan).Value2
        if ( $null -eq $cel_FAN ) {
            break
        }
        $cel_Attendance = $sh_Deadlines.Cells.Item($i, $col_att).Value2
        $b_Attendance = $true
        if ( $cel_Attendance -eq 0 ) {
            $b_Attendance = $false
        }

        # Add to our hash
        $h_AttendanceInfo[$cel_FAN] = $b_Attendance
        $i++ 
    }

    # Close the workbook
    $wb_Deadlines.Close($false)
    $excel.Quit()
    Remove-Variable excel

    return $h_AttendanceInfo
}

function MJFLateSubmissionPenalty ( $SubmissionDueDate, $SubmissionDate, $maximumMarks, $penaltyDecimal ) {
<#
    .Synopsis
    Calculates lateness of submission and applies penalty if so.

    .Description
    UPDATE: Not using the 'Submitted Late' column.  Can't rememmber why. 
    I think this is because the late submission did not take extensions
    into account (only overrides).  But using the Last modified (submisison)'
    field means we can take this into account.  So, 'Last Modified (Submission)'
    is put into $SubmissionDate and I can't rremember how we manage SubmissionDueDate.
    Oh yes! I add an extra column!  Because the piece o' crap FLO doesn't
    include a due date in the *individual* assessment reports even though
    it includes it in the *group* assessment reports!
    
    This is how I used to do it but I don't do it this way any more:

    FLO records late submissions in a the 'Submitted Late' column.  This
    column contains the number of days or number of hours submitted late
    eg: '2 days 4 hours late' or '44 mins 37 seconds late'.
    We use these entries to calculate late submissions.  As late submissions
    are calculated per day or part thereof then 'n days m hours' means
    that n+1 days are used for penalty purposes.  An entry of 'n mins m seconds'
    implies that the submisison is less than a day late and therefore a one
    day penalty is applied.

    .Parameter SubmissionDueDate
    String from 'Submitted late' field of FLO csv export

    .Parameter SubmissionDate
    String from 'Last Modified (submission) field of FLO csv export

    .Parameter maximumMarks
    Maximum marks attainable for this assessment

    .Parameter penaltyDecimal
    Late penalty expressed as a decimal (eg: 5% == 0.05)
#>

    # Convert the date strings to datetimes
    $d_currentSubmission = MJFConvertToDate -dateString $SubmissionDate
    # ...and get the deadline
    $d_submissionDeadline = MJFConvertToDate -dateString $SubmissionDueDate
    # Calculate submission lateness
    $d_submissionLateness = $d_submissionDeadline - $d_currentSubmission

    
    if ( $d_submissionLateness.TotalHours -lt 0 ) {
        # We reduce the mark by the specified percentage per day
        $f_reduction = [math]::Round(($d_submissionLateness.TotalHours / 24 * -1 ) + 0.5) * $maximumMarks * $penaltyDecimal
    } else {
        $f_reduction = 0.0
    }

    if ( $f_reduction -gt 0 ) {
        $lateFeedback = "Late penalty applied - submission deadline: $SubmissionDueDate submitted: $SubmissionDate"
    }

    # Return
    $f_reduction
    $lateFeedback
}

function MJFWordCountPenalty ( $penaltyDecimal, $penaltyWordIncrement, $wordLimit, $maximumMarks, $submissionFilePath ) {
<#
    .Synopsis
    Calculates word count of submission and applies penalty if over the limit.

    .Description
    Opens the MS Word file in $submissionFilePath and uses the Word word count
    feature to get the word count (excluding footnotes).  If the word count is
    over $wordLimit, then $penaltyDecimal of $maximumMarks is deducted from the 
    current mark given for each set of $penaltyWordIncrement words over the limit.

    .Parameter penaltyDecimal
    Penalty percentage expressed as decimail (eg: 5% == 0.05)

    .Parameter penaltyWordIncrement
    Number of words as increment to step penalty (eg: 5% per 100 words)

    .Parameter wordLimit
    The word limit for the paper

    .Parameter maximumMarks
    Maximum marks that can be attained for this paper

    .Parameter submissionFilePath
    Path to MS Word file that contains the paper
#>
    # Feedback always starts with this
    $lateFeedback = ""

    # Get the word count from the document
    if ( $submissionFilePath ) {
        #TODO: exception handling when opening non-existent or invalid documents
        try {
            $wordDoc = $MSWord.Documents.Open($submissionFilePath)
            Start-Sleep -Seconds 5
            $wordCount = $MSWord.ActiveDocument.ComputeStatistics(0, $false)
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "Error opening document: $errorMessage $submissionFilePath"
            $lateFeedback = "$lateFeedback Error opening document: $errorMessage $submissionFilePath"
        }
        finally {
            $wordDoc.close() | Out-Null
            Start-Sleep -Seconds 5
        }
    } else {
        $lateFeedback = "No submission file"
    }

    # TODO: What if word limit makes no sense?


    # Calculate the penalty
    if ( $wordCount -le $wordLimit ) {
        $wordPenalty = 0
    } else {
        # Calculate the penalty as the number of words over the limit divided by the increment
        $wordsOverLimit = $wordCount - $wordLimit
        $wordsOverIncrement = [math]::Round( ($wordsOverLimit / $penaltyWordIncrement) + 0.5)
        $wordPenalty = $wordsOverIncrement * $maximumMarks * $penaltyDecimal
        $lateFeedback = "$lateFeedback Word count: $wordCount  Limit: $wordLimit"
    }

    # Here's what to return
    $wordPenalty
    $lateFeedback
}

# MJFWriteFeedback
# Expect: marking record hash, late penalty string, word count penalty string, late penalty, word count penalty
# Return: Feedback string to insert into grades CSV
# This function needs to be customised to reflect each individual assessment
# as assessment criteria (and therefore object names) will vary from
# assessment to assessment.
function MJFWriteFeedback ( $currentRecord, $latePenalty, $latePenaltyString, $wcPenalty, $wcPenaltyString ) {

    $allFeedback = ""
    $currentGrade = $currentRecord.Grade
    $currentFeedback = $currentRecord.'Feedback comments'

    if ( $latePenalty + $wcPenalty -gt 0 ) {
        $netgrade = $currentGrade - $latePenalty - $wcPenalty
        $allFeedback = $allFeedback + "<p></p><p><b>Mark awarded (before penalties)</b> : $currentGrade</p>"
        if ( $latePenalty -gt 0 ) {
            $allFeedback = $allFeedback + "<p><b>Late penalty</b>      : $latePenalty ($latePenaltyString)</p>"
        }
        if ( $wcPenalty -gt 0 ) {
            $allFeedback = $allFeedback + "<p><b>Word count penalty</b>: $wcPenalty ($wcPenaltyString)</p>"
        }
    } else {
        $netgrade = $currentGrade
    }
    $netpercent = "{0:P0}" -f ($netgrade / $currentRecord.'Maximum Grade')
    $allFeedback = $allFeedback + "<p><b>Mark awarded:</b>           : $netgrade</p>"
    $allFeedback = $allFeedback + "<p><b>Mark as percentage:</b>     : $netpercent</p>"

    if ( $currentFeedback -ne "" ) {
        $allFeedback = $allFeedback + "<p>$currentFeedback</p>"
    }
    # Return the feedback
    $allFeedback
}


function main() {

    # Grades CSV (exported from FLO)
    $FLO_filepath, $FLO_csvdata = MJFGetCSVFile -displayText "Where is the grades export from FLO?"

    # Attendance data as a hash
    $h_AttendanceInfo = MJFGetAttendanceData

    # Submissions
    $submissions = MJFGetSubmissionFilePaths

    # We need to open up an instance of Word for word counting.
    $MSWord = New-Object -ComObject word.Application
    $MSWord.visible = $true

    # Iterate through marking table and process each row as we go
    $i_index = 1    # We just use this to display a basic progress indicator
    foreach ( $row in $FLO_csvdata ) {
        # Store FAN for easy reference
        $s_currentFAN = $row.FAN

        # Write some output to the console
        Write-Host "Processing $s_currentFAN ( $i_index / " $FLO_csvdata.Length ")"

        $i_index++

        # 'No submission' entries should be graded as 0
        if (  $row.'Individual submission status' -match '^No submission.*' ) {
            $row.Grade = 0
            $row.'Feedback comments' = "$row.'Feedback comments' <p>No submission</p>"
            Write-Host "$s_currentFAN - No submission"
            continue
        }

        # If the student didn't attend they should be graded as a 0
        if ( $null -ne $h_AttendanceInfo[$s_currentFAN] -and -not $h_AttendanceInfo[$s_currentFAN]) {
            $row.Grade = 0
            $row.'Feedback comments' = "Did not attend workshop."
            continue
        }

        # Overword penalty
        $wc_penalty, $wc_feedback = MJFWordCountPenalty -penaltyDecimal $overword_penalty -penaltyWordIncrement $overword_increment `
                                                        -wordLimit $overword_limit -maximumMarks $row.'Maximum Grade' -submissionFilePath $submissions[$s_currentFAN]

        

        # Apply Feedback
        $row.'Feedback comments' = MJFWriteFeedback -currentRecord $row -latePenalty $late_penalty -latePenaltyString $late_feedback `
                                        -wcPenalty $wc_penalty -wcPenaltyString $wc_feedback
        
        # Apply penalties to grade
        # Must happen after feedback
        $row.Grade -= $wc_penalty
        # Can't have a score less than zero
        if ( $row.Grade -lt 0 ) {
            $row.Grade = 0
        }

    }

    # Write the CSV
    $outputFile = $FLO_filepath + "_new.csv"
    $FLO_csvdata | Export-CSV -Path $outputFile -NoTypeInformation

    # Clean up
    $MSWord.Quit()
    Remove-Variable MSWord
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}

main