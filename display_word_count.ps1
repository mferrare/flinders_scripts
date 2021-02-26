# Displays word count of files in a folder

Import-Module .\MJFutils.psm1

# Where the files are
$dir_path = MJFGetDirPath h_Properties @{ Title = "Select folder containing assignment submissions" }

# Open connection to Word
# We need to open up an instance of Word for word counting.
$MSWord = New-Object -ComObject word.Application
$MSWord.visible = $true

# List out and process
Get-ChildItem $dir_path -Filter *.docx | ForEach-Object {
    $a_file = $_.FullName
    $a_name = $_.Name

    # Open the word file
    try {
        $wordCount = -1
        $wordDoc = $MSWord.Documents.Open($a_file)
        $wordCount = $wordDoc.ComputeStatistics(0, $false)
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error opening document: $a_name $errorMessage"
    }
    finally {
        try {
            $wordDoc.Close() | Out-Null
        }
        catch {
            # Doc didn't close.  Sleep and try again
            # That's as sophisticated as we get.
            Start-Sleep -Seconds 5
            $wordDoc.Close() | Out-Null
        }
    }
    Write-Host $wordCount, $a_name
}


# Close connection to Word and tidy up
$MSWord.Quit()
Remove-Variable MSWord
[gc]::collect()
[gc]::WaitForPendingFinalizers()