# I used this for my assignment submissions.  All submissions are
# downloaded into a single folder.  I iterate through each word file
# and convert it to PDF and store the PDFs in a folder called ../PDFs
Add-Type -AssemblyName System.Windows.Forms

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

# MJFMakeDir
# Expect: Path to parent of new directory to make, name of new directory
# Return: the path to the new directory or false
# Returns the path if the directory exists and is a directory or
# if the directory is created.  False otherwise.
function MJFMakeDir ( $s_dirPath, $s_newDir ){

    # Make the path to PDFs
    $parentPath = Split-Path -Path $s_dirPath -Parent


    $PDFpath = Join-Path -Path $parentPath -ChildPath $s_newDir

    # Check if it's a directory first
    if ( Test-Path -Path $PDFpath -PathType Container ) {
        return $PDFpath
    }

    # If we're here it either doesn't exist or isn't a directory.  If it exists and
    # is a file then return $false.
    if ( Test-Path -Path $PDFpath -PathType Leaf ) {
        return $false
    }

    # If we're here then the path doesn't exist.  Create it
    $null = New-Item -Path $parentPath -Name $s_newDir -ItemType "directory"

    # Test again to make sure it's there
    # Check if it's a directory first
    if ( Test-Path -Path $PDFpath -PathType Container ) {
        # Returning $PDFpath is also true
        return $PDFpath
    } else {
        return $false
    }
}

# MJFConvertWordFilesToPDF
# Expect: Path to folder containing word files, path to folder to store PDF files
# Return: Nothing
# Iterates through the folder looking for all Word files and uses Word's Save As
# item to save each file as a PDF
function MJFConvertWordFilesToPDF ($pathToWordFiles, $pathToPDFFiles) {

    $app_Word = New-Object -ComObject Word.Application

    # This filter will find .doc as well as .docx documents
    Get-ChildItem -Path $pathToWordFiles -Filter *.doc? | ForEach-Object {

        $document = $app_Word.Documents.Open($_.FullName)
        $pdfFilename = $_.BaseName + ".pdf"
        $pdfFileFullPath = Join-Path -Path $pathToPDFFiles -ChildPath $pdfFilename
        $document.SaveAs([ref][system.object] $pdfFileFullPath, [ref] 17)
        $document.Close()
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
    $submissionPDFDir = MJFMakeDir -s_dirPath $submissionDir -s_newDir "PDF"

    # Only proceed if we have successfully made the directory
    if ( $submissionPDFDir ){
        # Iterate through the word files and convert them to PDF
        MJFConvertWordFilesToPDF -pathToWordFiles $submissionDir -pathToPDFFiles $submissionPDFDir
    }
}

__main__