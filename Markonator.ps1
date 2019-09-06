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

# MJFGetDirPath
# Returns the path of the selected directory
function MJFGetDirPath ( $h_Properties ){

    $DirBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        ShowNewFolderButton = $false
    }

    foreach ( $key in $h_Properties.keys ) {
        $DirBrowser.$key = $h_Properties.$key
    }
    $null = $DirBrowser.ShowDialog()
    
    return $DirBrowser.SelectedPath
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

# MJFGetAttendanceData
# Expect: Nothing
# Return: A hash indexed by FAN with this information:
# - whether they attended the latest seminar
# - Their seminar number (to calculate submission lateness)
# - Their workshop group number (actaully, not sure I need this!  But I've written the code already so...)
function MJFGetAttendanceData() {
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
    $cel_Start = $sh_Deadlines.Cells.Item(4, 1)
    # TODO: Check the cell value contains Student ID.  Otherwise we're reading the wrong spreadsheet
    $null = $cel_Start.Value2

    # Read through the data in the table
    # We assume that the first empty row signifies the end of the table
    $h_AttendanceInfo = @{ }
    $i = 5
    while ($true) {
        $cel_FAN = $sh_Deadlines.Cells.Item($i, 2).Value2
        if ( $null -eq $cel_FAN ) {
            break
        }
        $cel_Attendance = $sh_Deadlines.Cells.Item($i, 17).Value2
        $b_Attendance = $true
        if ( $cel_Attendance -eq 0 ) {
            $b_Attendance = $false
        }

        # Find which seminar group and workshop group we're in
        $s_SeminarNumber = $null
        $s_WorkshopGroup = $null
        $a_temp = ($sh_Deadlines.cells.Item($i, 6).Value2).split(',')
        foreach ( $element in $a_temp) {
            # We need to examine our array of seminar groups and etxtract:
            # - the workshop we're in (ie: 'Seminar 0n')
            # - the workshop group we're in (ie: 'S0nG0n')
            # Pattern matching to the rescue!
            $s_temp = $element.trim()

            if ( $s_temp -match 'Seminar [0-9][0-9]' ) {
                $s_SeminarNumber = $s_temp
            }
            if ( $s_temp -match 'S[0-9][0-9]G[0-9][0-9]' ) {
                $s_WorkshopGroup = $s_temp
            }
        }
        # Add everything to our hash

        $h_AttendanceInfo[$cel_FAN] = @{ Attendance = $b_Attendance; SeminarNumber = $s_SeminarNumber; WorkshopGroup = $s_WorkshopGroup }
        $i++ 
    }

    # Close the workbook
    $wb_Deadlines.Close($false)
    $excel.Quit()
    Remove-Variable excel

    return $h_AttendanceInfo
}

# MJFGetMarkingSpreadsheet
# Expect: Nothing
# Return: path to CSV, CSV data imported into table
# Opens and imports the marking spreadsheet.
# The marking 'spreadsheet' is a CSV export from FLO
function MJFGetMarkingSpreadsheet() {
    $fp_MarkingCSV = MJFGetFilePath -h_Properties @{
        Title = "Where is the Marking Spreadsheet?"
        Filter = 'CSV File (*.csv)|*.csv'
    }
    #$fp_MarkingCSV = "C:\Users\ferr0182\OneDrive - Flinders\LLAW2221_2019_S2\Workshop Assignments\Workshop 2\LLAW2221_2019_S2_Written Assignment - Workshop 2_grades - Test.csv"

    $t_MarkingTable = Import-Csv -Path $fp_MarkingCSV

    $fp_MarkingCSV
    $t_MarkingTable
}

# MJFGetSubmissionFilePaths
# Iterate through the folder containing submissions.  Create a hash with
# key   : FAN
# value : Full path to word document
function MJFGetSubmissionFilePaths() {

    # return result in this hash
    $result = @{}

    # Get the folder where the documents are stored
    $WordDocFolder = MJFGetDirPath -h_Properties @{ Description = "Select folder containing assignment submissions" }
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
        $wordDoc = $MSWord.Documents.Open($submissionFilePath)
        $wordCount = $MSWord.ActiveDocument.ComputeStatistics(0, $false)
        $wordDoc.close() | Out-Null
        #[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
        #Remove-Variable wordDoc
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
        $lateFeedback = "Over word count penalty: " + $wordPenalty + " marks."
    }

    # Here's what to return
    $wordPenalty
    $lateFeedback
}

#MJFLateSubmissionPenalty
# Expect: Seminar number, submission DateTime, max marks attainable, penalty to apply to max marks
# Return: array [ (float) penalty, (string) feedback ]
# Calculates lateness of submission and applies penalty if so.
function MJFLateSubmissionPenalty ( $SeminarNumber, $SubmissionDateTime, $maximumMarks, $penaltyDecimal ) {

    # Feedback to return
    $lateFeedback = ""
    if ( $null -eq $SeminarNumber ) {
        $lateFeedback += " No Seminar number found."
        $f_reduction = 0
        $f_reduction
        $lateFeedback
        return
    }

    # Otherwise, process away!
    # Process away!
    $s_DeadlineTime = $h_Deadlines[$SeminarNumber]['Time']
    $s_DeadlineDay = $h_Deadlines[$SeminarNumber]['Day']
    # Get the date...
    $d_DeadlineDateTime = Get-Date
    for ( $i = -1; $i -gt -7; $i-- ) {
        if ( $d_DeadlineDateTime.AddDays($i).DayOfWeek -eq $s_DeadlineDay ) {
            $d_DeadlineDateTime = $d_DeadlineDateTime.AddDays($i)
            break
        }
    }
    # ...and the time
    $hour,$min = $s_DeadlineTime.split(':')

    # Set the Deadline as a DateTime
    $d_DeadlineDateTime = Get-Date -Date $d_DeadlineDateTime -Hour $hour -Minute $min -Second 0

    # Calculate submission lateness
    $d_submissionLateness = $d_DeadlineDateTime - $SubmissionDateTime
    if ( $d_submissionLateness.TotalHours -lt 0 ) {
        # We reduce the mark by 5% per day
        # TODO.  >= 24 hours needs to be rounded up to the nearest 24 hours
        $f_reduction = [math]::Round(($d_submissionLateness.TotalHours / 24 * -1 ) + 0.5) * $maximumMarks * $penaltyDecimal
        $lateFeedback += "Late submission penalty applied: " + $f_reduction + " marks."
    } else {
        $f_reduction = 0.0
    }

    # Return
    $f_reduction
    $lateFeedback
}

#
# Main Program
# 

# Get our data
$h_Deadlines = MJFGetSubmissionDeadlines
$h_AttendanceInfo = MJFGetAttendanceData
$s_markingFilePath, $a_MarkingTable = MJFGetMarkingSpreadsheet
$h_SubmissionFileNames = MJFGetSubmissionFilePaths

# We need to open up an instance of Word for word counting.
$MSWord = New-Object -ComObject word.Application
$MSWord.visible = $false

# Iterate through marking table and process each row as we go
foreach ( $row in $a_MarkingTable) {
    # Store FAN for easy reference
    $s_currentFAN = $row.FAN
  
    # Before we get any further, make sure we have an attendance record for the FAN
    # If we don't, there's no point processing as that FAN hasn't been allocated
    # a group (as opposed to being absent)
    if ( $null -eq $h_AttendanceInfo[$s_currentFAN] ) {
        continue
    }

    # Check attendance.  If the FAN didn't attend the seminar then set the
    # grade to 0 and continue
    if ( $null -ne $h_AttendanceInfo[$s_currentFAN]['Attendance'] -and -not $h_AttendanceInfo[$s_currentFAN]['Attendance']) {
        $row.Grade = 0
        $row.'Feedback comments' = "Did not attend workshop."
        continue
    }

    # Easy reference to seminar number
    $s_currentSeminarNumber = $h_AttendanceInfo[$s_currentFAN]['SeminarNumber']
    # Exit if no seminar number.  Means not enrolled.  Set grade to 0 just in case.
    if ( $null -eq $s_currentSeminarNumber ) {
        $row.Grade = 0
        $row.'Feedback comments' = "Not enrolled in any seminar"
        continue
    }

    # Store feedback in here
    $row_feedback = ""

    # Late penalty calculation
    # Convert the date in the marking CSV to a dateTime
    $s_currentSubmission = $row.'Last modified (submission)'
    $d_currentSubmission = [dateTime] $s_currentSubmission.split(',',2)[1].trim()
    # get the penalty and apply it
    $latePenalty, $lateFeedback = MJFLateSubmissionPenalty -SeminarNumber $s_currentSeminarNumber `
                                    -SubmissionDateTime $d_currentSubmission `
                                    -maximumMarks $row.'Maximum Grade' `
                                    -penaltyDecimal 0.05 
    $row.Grade -= $latePenalty
    
    $row_feedback += " " + $lateFeedback

    # Word count
    $wordPenalty, $wcFeedback = MJFWordCountPenalty -penaltyDecimal 0.05 `
         -wordLimit $h_Deadlines[$s_currentSeminarNumber]['WordCount'] `
         -submissionFilePath $h_SubmissionFileNames[$s_currentFAN] `
         -maximumMarks $row.'Maximum Grade' `
         -penaltyWordIncrement 100
    $row.Grade -= $wordPenalty
    $row_feedback += " " + $wcFeedback

    # Write feedback
    $row.'Feedback comments' = $row_feedback.Trim()
}

# Write the CSV
$outputFile = $s_markingFilePath + "_new.csv"
$a_MarkingTable | Export-CSV -Path $outputFile -NoTypeInformation

# Clean up
$MSWord.Quit()
Remove-Variable MSWord
[gc]::collect()
[gc]::WaitForPendingFinalizers()