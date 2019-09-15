Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# A class to store the group of widgets that
class FileLoader {
    # Constants
    $VPADDING = 3
    $HPADDING = 3
    $TEXTBOXWIDTH = 300
    # TODO: Calculate height from text box font
    $TEXTBOXHEIGHT = 20

    # Form Elements
    $titleLabel
    $filePathTextBox

    # Constructor
    FileLoader([String]$titleString)
    {
        $this.titleLabel = New-Object System.Windows.Forms.Label
        $this.titleLabel.Text = $titleString
        $this.titleLabel.AutoSize = $true
        $this.titleLabel.BorderStyle = 2

        $this.filePathTextBox = New-Object System.Windows.Forms.TextBox
        $this.filePathTextBox.ScrollBars = 1 # Horizontal only (ScrollBars.Horizontal)
        $this.filePath.TextBox.ClientSize = New-Object System.Drawing.Size($this.TEXTBOXWIDTH, $this.TEXTBOXHEIGHT)
    }

    # Eventually, this method will display all the objects in  this 
    # class.  All that should need to be supplied is the x,y co-ordinates
    # to position the titleLabel.  All other objects will be positioned 
    # relative to that.
    displayAllObjects([int]$x, [int]$y, $theForm) {
        # cursor_x and cursor_y we use to track where to put the next element
        # Start with what we receive
        $cursor_x = $x
        $cursor_y = $y

        $this.titleLabel.Location = New-Object System.Drawing.Point($cursor_x, $cursor_y)
        $theForm.Controls.Add($this.titleLabel)
        $cursor_y += $this.titleLabel.Location.
        $this.filePathTextBox.Location
    }
}

# Main Program
$theForm = New-Object System.Windows.Forms.Form
$theForm.AutoSize = $true
$theForm.StartPosition = 'CenterScreen'

$deadlinesLoader = [FileLoader]::new("Deadlines File")
$deadlinesLoader.displayAllObjects(10, 10, $theForm)

$theForm.Show()