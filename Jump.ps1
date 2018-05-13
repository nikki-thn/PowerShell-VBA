
#create window form
$Form = New-Object system.Windows.Forms.Form
$Form.Width = 478
$Form.Height = 280
$form.MaximizeBox = $false 
$Form.StartPosition = "CenterScreen" 
$Form.Text = "Quick Open"

#Add button to the form
$button00 = New-Object System.Windows.Forms.Button 
$button00.Location = New-Object System.Drawing.Size(0,0) 
$button00.Size = New-Object System.Drawing.Size(100,30) 
$button00.Text = "Model.docx" 
$button00.BackColor = "LightBlue"
$button00.AutoSize = $true 
$Form.Controls.Add($button00)

#Add button to the form
$button10 = New-Object System.Windows.Forms.Button 
$button10.Location = New-Object System.Drawing.Size(0,40) 
$button10.Size = New-Object System.Drawing.Size(100,30) 
$button10.Text = "Modelset.docx"
$button10.BackColor = "LightBlue" 
$button10.AutoSize = $true 
$Form.Controls.Add($button10)

#Add button to the form
$button20 = New-Object System.Windows.Forms.Button 
$button20.Location = New-Object System.Drawing.Size(0,80) 
$button20.Size = New-Object System.Drawing.Size(100,30) 
$button20.Text = "THH_Models" 
$button20.BackColor = "LightBlue" 
$button20.AutoSize = $true 
$Form.Controls.Add($button20)

#Add button to the form
$button30 = New-Object System.Windows.Forms.Button 
$button30.Location = New-Object System.Drawing.Size(0,120) 
$button30.Size = New-Object System.Drawing.Size(100,30) 
$button30.Text = "THH_Modelsets"
$button30.BackColor = "LightBlue" 
$button30.AutoSize = $true 
$Form.Controls.Add($button30)

#Add button to the form
$button40 = New-Object System.Windows.Forms.Button 
$button40.Location = New-Object System.Drawing.Size(0,160) 
$button40.Size = New-Object System.Drawing.Size(100,30) 
$button40.Text = "DOCS DEV_Sup" 
$button40.BackColor = "Gold" 
$button40.AutoSize = $true 
$Form.Controls.Add($button40)

#Add button to the form
$button50 = New-Object System.Windows.Forms.Button 
$button50.Location = New-Object System.Drawing.Size(0,200) 
$button50.Size = New-Object System.Drawing.Size(100,30) 
$button50.Text = "Archive" 
$button50.BackColor = "Gold" 
$button50.AutoSize = $true 
$Form.Controls.Add($button50)

#Add button to the form
$button01 = New-Object System.Windows.Forms.Button 
$button01.Location = New-Object System.Drawing.Size(120,0) 
$button01.Size = New-Object System.Drawing.Size(100,30) 
$button01.Text = "CR Templates"
$button01.BackColor = "Gold"  
$button01.AutoSize = $true 
$Form.Controls.Add($button01)

#Add button to the form
$button02 = New-Object System.Windows.Forms.Button 
$button02.Location = New-Object System.Drawing.Size(240,0) 
$button02.Size = New-Object System.Drawing.Size(100,30) 
$button02.Text = "BB Templates" 
$button02.BackColor = "Gold" 
$button02.AutoSize = $true 
$Form.Controls.Add($button02)

#Add button to the form
$button03 = New-Object System.Windows.Forms.Button 
$button03.Location = New-Object System.Drawing.Size(360,0) 
$button03.Size = New-Object System.Drawing.Size(100,30) 
$button03.Text = "AMDR DEV" 
$button03.BackColor = "Gold" 
$button03.AutoSize = $true 
$Form.Controls.Add($button03)

#Add button to the form
$button11 = New-Object System.Windows.Forms.Button 
$button11.Location = New-Object System.Drawing.Size(120,40) 
$button11.Size = New-Object System.Drawing.Size(100,30) 
$button11.Text = "H:" 
$button11.BackColor = "Salmon" 
$button11.AutoSize = $true 
$Form.Controls.Add($button11)

#Add button to the form
$button12 = New-Object System.Windows.Forms.Button 
$button12.Location = New-Object System.Drawing.Size(240,40) 
$button12.Size = New-Object System.Drawing.Size(100,30) 
$button12.Text = "Feed Files" 
$button12.BackColor = "Gold" 
$button12.AutoSize = $true 
$Form.Controls.Add($button12)

#Add button to the form
$button13 = New-Object System.Windows.Forms.Button 
$button13.Location = New-Object System.Drawing.Size(360,40) 
$button13.Size = New-Object System.Drawing.Size(100,30) 
$button13.Text = "Empty" 
$button13.AutoSize = $true 
$Form.Controls.Add($button13)

#Add button to the form
$button14 = New-Object System.Windows.Forms.Button 
$button14.Location = New-Object System.Drawing.Size(480,40) 
$button14.Size = New-Object System.Drawing.Size(100,30) 
$button14.Text = "Empty" 
$button14.AutoSize = $true 
$Form.Controls.Add($button14)

#ALIAS FUNCTIONS
#Please replace with alias functions

#Event Handler Button on-click
$button00.Add_Click({Open-Model})
$button10.Add_Click({Open-Modelset})
$button20.Add_Click({Open-THH-Models})
$button30.Add_Click({Open-THH-Modelsets})
$button40.Add_Click({Open-DEV})
$button50.Add_Click({Open-Archive})

$button01.Add_Click({Open-CRTemplates})
$button02.Add_Click({Open-BBTemplates})
$button03.Add_Click({Open-AMDR-DEV})

$button11.Add_Click({Open-H})
$button12.Add_Click({Open-Feedfiles})

#Open the form
$Form.ShowDialog()
