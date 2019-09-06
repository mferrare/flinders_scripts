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

# Let's create a label
# Text to hash goes in here
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Deadlines File:'
$label.Location = New-Object System.Drawing.Point($posLeftMostSide,$posSingleLineHeight)
$label.Size = New-Object System.Drawing.Size($posSingleLabelWidth,$posSingleLineHeight)