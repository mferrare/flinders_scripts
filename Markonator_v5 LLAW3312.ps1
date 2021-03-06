﻿# - Reads grade from a FLO CSV file.
# - Looks for submission word file and counts words and
#   adjusts penalties accordingly.
# - Reads submission times from CSV file and deducts
#   late marks accordingly.
#
# *** IMPORTANT ***
# This script depends on the existence of a 'Due Date' column in the FLO
# report.  This column *does not exist* for reports generated for individual
# assessments but does exist for reports generated for group assessements (don't ask)
# If the column doesn't exist you must:
# - Add a 'Due Date' column
# - Fill it with the correct due date for each student.  Make sure you
#   account for extensions
# The due date format should be the FLO format ie:
#   Monday, 24 August 2020 10:00 AM'

Import-Module .\MJFutils.psm1

# Global Variables
$late_submission_penalty = 0.05  # % penalty as a decimal
$overword_limit = 2000
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

function MJFLateSubmissionPenaltyWithDueDate ( $SubmissionDueDate, $SubmissionDate, $maximumMarks, $penaltyDecimal ) {
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

function MJFLateSubmissionPenaltyWithLateSubmission ( $late_submission, $maximumMarks, $penaltyDecimal ) {
<#
    .Synopsis
    Calculates lateness of submission and applies penalty if so.

    .Description
    FLO records late submissions in a the 'Submitted Late' column.  This
    column contains the number of days or number of hours submitted late
    eg: '2 days 4 hours late' or '44 mins 37 seconds late'.
    We use these entries to calculate late submissions.  As late submissions
    are calculated per day or part thereof then 'n days m hours' means
    that n+1 days are used for penalty purposes.  An entry of 'n mins m seconds'
    implies that the submisison is less than a day late and therefore a one
    day penalty is applied.

    .Parameter late_submission
    String from 'Submitted late' field of FLO csv export

    .Parameter maximumMarks
    Maximum marks attainable for this assessment

    .Parameter penaltyDecimal
    Late penalty expressed as a decimal (eg: 5% == 0.05)
#>

    $days_late = 0 # Store number of days late here.

    # Only test if there is something in $late_submission
    if ( $late_submission ) {

        # Assume that if 'days' does not exist in the string then 
        # submission is less than one day late.
            
        $days_late_match = Select-String -InputObject $late_submission -Pattern '([0-9]+) day.*late'
        if ( $days_late_match.Matches ) {
            # Get the number of days
            $days_late = [int]$days_late_match.Matches[0].Groups[1].Value

            # Add one to days_late if we're not right on a day boundary (eg: '2 days late' vs '2 days 2 hours late')
            # Look for the day boundary
            $just_a_day = Select-String -InputObject $late_submission -Pattern '^([0-9]+) days{0,1} late$'
            if ( -not $just_a_day.Matches.Success ) {
                # And increment if we're not on the boundary
                $days_late = $days_late + 1
            }
        } else {
            $days_late = 1
        }
    }

    $f_reduction = $days_late * $maximumMarks * $penaltyDecimal
    if ( $f_reduction -gt 0 ) {
        $lateFeedback = "Late penalty applied.  Submitted $late_submission"
    }

    # Return
    $f_reduction
    $lateFeedback
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

    # Get the Rubric Criteria
    $c1 = $currentRecord.'Criterion 1'
    $c2 = $currentRecord.'Criterion 2'
    $c3 = $currentRecord.'Criterion 3'
    $c4 = $currentRecord.'Criterion 4'
    $c5 = $currentRecord.'Criterion 5'
    
    $allFeedback = ""
    $currentGrade = $currentRecord.Grade
    $currentFeedback = $currentRecord.'Feedback comments'

    # Add rubric criteria to feedback
    $allFeedback = "<p><b>Criterion 1 (Research)</b>                  : $c1</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 2 (Analysis)</b>   : $c2</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 3 (Structure)</b>  : $c3</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 4 (AGLC)</b>       : $c4</p>"
    $allFeedback = $allFeedback + "<p><b>Criterion 5 (Readability)</b>: $c5</p>"

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

    # Grades Spreadsheet
    $FLO_filepath, $FLO_csvdata = MJFGetCSVFile -displayText "Where is the grades export from FLO?"

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
        Write-Host "Processing $s_currentFAN ( $i_index /"$FLO_csvdata.Length")"

        $i_index++

        # 'No submission' entries should be graded as 0
        if (  $row.'Individual submission status' -match '^No submission.*' ) {
            $row.Grade = 0
            $row.'Feedback comments' = "$row.'Feedback comments' <p>No submission</p>"
            Write-Host "$s_currentFAN - No submission"
            continue
        }

        # Late submission penalty
        $late_penalty, $late_feedback = MJFLateSubmissionPenaltyWithLateSubmission -late_submission $row.'Submitted late' `
                                                                 -maximumMarks $row.'Maximum Grade' `
                                                                 -penaltyDecimal $late_submission_penalty

        # Overword penalty
        $wc_penalty, $wc_feedback = MJFWordCountPenalty -penaltyDecimal $overword_penalty -penaltyWordIncrement $overword_increment `
                                                        -wordLimit $overword_limit -maximumMarks $row.'Maximum Grade' -submissionFilePath $submissions[$s_currentFAN]

        

        # Apply Feedback
        $row.'Feedback comments' = MJFWriteFeedback -currentRecord $row -latePenalty $late_penalty -latePenaltyString $late_feedback `
                                        -wcPenalty $wc_penalty -wcPenaltyString $wc_feedback
        
        # Apply penalties to grade
        # Must happen after feedback
        $row.Grade -= $late_penalty
        $row.Grade -= $wc_penalty
        # Can't have a score less than zero
        if ( $row.Grade -lt 0 ) {
            $row.Grade = 0
        }

    }

    # Write the feedback
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