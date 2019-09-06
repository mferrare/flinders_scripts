
#Here is an example how to call the powershell function which displays a window with any image. 
#When you click on a pixel in the image it displays the parts of Red, Green and Blue for that pixel.  #You can use this to help identify which color spectrums you want to remove or update when cleaning up images for OCR'ing your images.
#get-ImagePixelColor "$home\Pictures\PayStubSinglePage.jpg"
 
Function get-ImagePixelColor{
    [CmdletBinding()]
    Param(  [Parameter(Mandatory=$True,Position=1)] [string]$FileName)

    [void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
    $file = (get-item $FileName)
    $img = [System.Drawing.Image]::Fromfile($file);
 
    $form = new-object Windows.Forms.Form
    $form.Text = "Display Image"
    $form.Width = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width-50;
    $form.Height = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height-50;

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Height =  $form.Height-50;
    $panel.Width = $form.Width-200;
    $panel.AutoScroll = $true

    function createLabel {
      Param(  [Parameter(Mandatory=$True)] [int]$Top,
              [Parameter(Mandatory=$True)] [System.Drawing.Color]$BackColor)     
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Height = 50
        $lbl.Width = 100
        $lbl.Left = $form.Width-150;
        $lbl.Top = $top;
        $lbl.Text = ""
        $lbl.BackColor = $BackColor
        $font = New-Object System.Drawing.Font([System.Drawing.FontFamily]::GenericSansSerif, 24)
        $lbl.Font = $font    
        return $lbl
    }
    $lblBlue = createLabel 50 $([System.Drawing.Color]::Blue)
    $lblRed  = createLabel 125 $([System.Drawing.Color]::Red)
    $lblGreen = createLabel 200 $([System.Drawing.Color]::Green)

    #Mouse click event for pictureBox
    function pictureBox_MouseClick($Sender, $EventArgs)
    { 
        $image = new-object System.Drawing.Bitmap $pictureBox.Image;
        $color = $image.GetPixel($EventArgs.X, $EventArgs.Y);
        $lblBlue.Text = $color.B
        $lblGreen.Text = $color.G
        $lblRed.Text = $color.R
    }
    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.SizeMode =  [System.Windows.Forms.PictureBoxSizeMode]::AutoSize;   
    $pictureBox.Image = $img;
    $pictureBox.add_MouseClick({pictureBox_MouseClick $this $_})

    $form.controls.add($lblBlue)
    $form.controls.add($lblRed)
    $form.controls.add($lblGreen)
    $form.controls.add($panel)
    $panel.controls.add($pictureBox)
    $form.Add_Shown( { $form.Activate() } )
    $form.ShowDialog()
}
get-ImagePixelColor
