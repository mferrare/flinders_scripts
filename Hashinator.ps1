Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Click event
$Button_Click = 
{
    $hashedString = doHashOfString($textBox.Text)
    $hashLabel.Text = $hashedString
}
# Mining button event
$MineButton_Click = 
{
    $hashLabel.Text = 'Working...'
    $hashLabel.Refresh()
    $textToHash = $textBox.Text
    $difficulty = $DifficultyTextBox.Text
    doMining -stringToMine $textToHash -numberOfZeroes $difficulty
}
function doHashOfString([String] $stringToHash)
{
    $hasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256')
    $hash = $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToHash))
    $hashString = [System.BitConverter]::ToString($hash)
    $hashString.Replace('-', '')
    return 
}

function doMining
{
    Param($stringToMine, $numberOfZeroes)

    # The most zeroes we could ever mine
    $allTheZeroes = "000000000000000000000000000000000000"

    $zeroCheck = $allTheZeroes.Substring(0,$numberOfZeroes)

    # Start at $i = 0
    $i = 0

    while ($true)
    {
        $NonceTextBox.Text = $i
        $NonceTextBox.Refresh()
        # Concatenate $stringToMine to $numberOfZeroes
        $stringToHash = "$stringToMine $i"

        # Hash the catenated string
        $hashedString = doHashOfString($stringToHash)

        # Check for $numberOfZeroes zeroes
        $leftSide = $hashedString.SubString(0,$numberOfZeroes)

        # Exit the loop if we have the correct number of zeroes
        if ($leftSide -eq $zeroCheck) {
            $hashLabel.Text = $hashedString
            break
        }
        $i++
    }
}

# Positioning variables
$posFormWidth = 800
$posFormHeight = 600
$posLeftMostSide = 10
$posSeparationBufferX = 10
$posSeparationBufferY = 10
$posSingleLineHeight = 30
$posSingleLabelWidth = 370
$posButtonStartX = 75
$posButtonStartY = 500
$posButtonHeight = 50
$posButtonWidth = 200
# Fonts to use
$standardFont = New-Object System.Drawing.Font("Microsoft Sans Serif",18,0,3,0)
$fixedFont = New-Object System.Drawing.Font("Consolas",24,0,3,0)

$form = New-Object System.Windows.Forms.Form
$form.Font = $standardFont
$form.Text = 'Data Entry Form'
$form.Size = New-Object System.Drawing.Size($posFormWidth, $posFormHeight)
$form.StartPosition = 'CenterScreen'

# Declare the buttons in order.  Make sure the X/Y co-ords all match up
# Hash button
$HashButton = New-Object System.Windows.Forms.Button
$HashButton.Location = New-Object System.Drawing.Point($posButtonStartX,$posButtonStartY)
$HashButton.Size = New-Object System.Drawing.Size($posButtonWidth,$posButtonHeight)      
$HashButton.Text = 'Hash me!'
$HashButton.Add_Click($Button_Click)
$form.Controls.Add($HashButton)
# Mine button
$MineButton = New-Object System.Windows.Forms.Button
$MineButtonX = $HashButton.Location.X + $HashButton.Size.Width
$MineButton.Location = New-Object System.Drawing.Point($MineButtonX,$posButtonStartY)
$MineButton.Size = New-Object System.Drawing.Size($posButtonWidth,$posButtonHeight)
$MineButton.Text = 'Mine me!'
$MineButton.Add_Click($MineButton_Click)
$form.Controls.Add($MineButton)
# Cancel Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButtonX = $MineButton.Location.X + $MineButton.Size.Width
$CancelButton.Location = New-Object System.Drawing.Point($CancelButtonX,$posButtonStartY)
$CancelButton.Size = New-Object System.Drawing.Size($posButtonWidth,$posButtonHeight)
$CancelButton.Text = 'Cancel :-('
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

# Text to hash goes in here
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Enter data to hash:'
$label.Location = New-Object System.Drawing.Point($posLeftMostSide,$posSingleLineHeight)
$label.Size = New-Object System.Drawing.Size($posSingleLabelWidth,$posSingleLineHeight)
$form.Controls.Add($label)
# Input
$textBox = New-Object System.Windows.Forms.TextBox
$textBoxX = $label.Location.X + $label.Size.Width
$textBox.Location = New-Object System.Drawing.Point($textBoxX,$posSingleLineHeight)
$textBox.Size = New-Object System.Drawing.Size(350,$posSingleLineHeight)
$form.Controls.Add($textBox)

# Result is displayed here
$hashLabel = New-Object System.Windows.Forms.Label
$hashLabel.Font = $fixedFont
$hashLabelY = $textBox.Location.Y + $textBox.Size.Height + $posSeparationBufferY
$hashLabel.Location = New-Object System.Drawing.Point($posLeftMostSide,$hashLabelY)
$hashLabel.Size = New-Object System.Drawing.Size(730,100)
$hashLabel.BorderStyle = 1
$form.Controls.Add($hashLabel)

# Mining fields.  A Place to set difficulty and a place to record the nonce
$Difficultylabel = New-Object System.Windows.Forms.Label
$Difficultylabel.Text = 'Set difficulty (number of zeroes):'
$DifficultyLabelY = $hashLabel.Location.Y + $hashLabel.Size.Height + $posSeparationBufferY
$Difficultylabel.Location = New-Object System.Drawing.Point($posLeftMostSide,$DifficultyLabelY)
$Difficultylabel.Size = New-Object System.Drawing.Size($posSingleLabelWidth,$posSingleLineHeight)
$form.Controls.Add($Difficultylabel)
# Input
$DifficultyTextBox = New-Object System.Windows.Forms.TextBox
$DifficultyTextBoxX = $DifficultyLabel.Location.X + $DifficultyLabel.Size.Width + $posSeparationBufferX
$DifficultyTextBoxY = $DifficultyLabel.Location.Y 
$DifficultyTextBox.Location = New-Object System.Drawing.Point($DifficultyTextBoxX,$DifficultyTextBoxY)
$DifficultyTextBox.Size = New-Object System.Drawing.Size(50,$posSingleLineHeight)
$form.Controls.Add($DifficultyTextBox)
# Nonce label
$NonceLabel = New-Object System.Windows.Forms.Label
$NonceLabel.Text = 'Nonce:'
$NonceLabelY = $DifficultyTextBox.Location.Y + $DifficultyTextBox.Size.Height + $posSeparationBufferY
$NonceLabel.Location = New-Object System.Drawing.Point($posLeftMostSide,$NonceLabelY)
$NonceLabel.Size = New-Object System.Drawing.Size($posSingleLabelWidth,$posSingleLineHeight)
$form.Controls.Add($NonceLabel)
# Input
$NonceTextBox = New-Object System.Windows.Forms.TextBox
$NonceTextBoxY = $NonceLabelY
$NonceTextBoxX = $NonceLabel.Location.X + $NonceLabel.Size.Width + $posSeparationBufferX
$NonceTextBox.Location = New-Object System.Drawing.Point($NonceTextBoxX,$NonceTextBoxY)         # Location
$NonceTextBox.Size = New-Object System.Drawing.Size(350,$posSingleLineHeight)              # Size
$form.Controls.Add($NonceTextBox)


$form.Topmost = $false

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}