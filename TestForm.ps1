 
     
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="MyForm" 
    $MyForm.Size = New-Object System.Drawing.Size(800,600) 
     
 
        $mTB1 = New-Object System.Windows.Forms.TextBox 
                $mTB1.Text="TextBox1" 
                $mTB1.Top="50" 
                $mTB1.Left="25" 
                $mTB1.Anchor="Left,Top" 
        $mTB1.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mTB1) 
        $MyForm.ShowDialog()
