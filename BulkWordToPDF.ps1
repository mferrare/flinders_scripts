# I used this for my assignment submissions.  All submissions are
# downloaded into a single folder.  I iterate through each word file
# and convert it to PDF and store the PDFs in a folder called ../PDFs
Add-Type -AssemblyName System.Windows.Forms

Import-Module .\MJFutils.psm1

# Global variables
$nameOfDirectory = "PDF"



# MJFConvertWordFilesToPDF
# Expect: Path to folder containing word files, path to folder to store PDF files
# Return: Nothing
# Iterates through the folder looking for all Word files and uses Word's Save As
# item to save each file as a PDF
function MJFConvertWordFilesToPDF ($pathToWordFiles, $pathToPDFFiles) {

    $app_Word = New-Object -ComObject Word.Application

    # This filter will find .doc as well as .docx documents
    $item_list = Get-ChildItem -Path $pathToWordFiles -Filter *.doc?

    # Store the count so we can show a basic progress indicator
    $num_items = $item_list.count

    # Process each file.
    $item_list | ForEach-Object { $i = 1 } {

        Write-Host "Trying $i / $num_items ($_.FullName)"

        $document = $app_Word.Documents.Open($_.FullName)
        $pdfFilename = $_.BaseName + ".pdf"
        $pdfFileFullPath = Join-Path -Path $pathToPDFFiles -ChildPath $pdfFilename
        $document.SaveAs([ref][system.object] $pdfFileFullPath, [ref] 17)
        $document.Close()

        $i += 1
    }

    $app_Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app_Word)
    Remove-Variable app_Word
}

# Main function.  We put the main program in a function so it's
# easier to find when scrubbing through the file.
function __main__() {
    # Find the submissions directory.
    # Find the CSV file with the deadlines...
    $submissionDir = MJFGetDirPath -h_Properties @{ 
        Title = "Select any Word document in the Submissions folder";
        Filter = 'Word Files (*.doc?)|*.doc?';
    }

    # We store the PDFs in a different directory at the same level
    $submissionPDFDir = MJFMakeDir -s_dirPath $submissionDir -s_newDir $nameOfDirectory

    # Only proceed if we have successfully made the directory
    if ( $submissionPDFDir ){
        # Iterate through the word files and convert them to PDF
        MJFConvertWordFilesToPDF -pathToWordFiles $submissionDir -pathToPDFFiles $submissionPDFDir
    }
}

__main__