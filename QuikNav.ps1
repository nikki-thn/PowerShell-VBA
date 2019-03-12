# QuikNav.ps1
# Version 1.0
# Windows widget that provides quick access to a folder/files or trigger a script
# Link: 
# Last modified on: May 25, 2018

Add-Type -AssemblyName System.Windows.Forms 

#create window form
$Form = New-Object system.Windows.Forms.Form
#set form size 
$Form.Width = 336 # width of the form
$Form.Height = 268 # height of the form
$form.MaximizeBox = $false #disable maximize window form
$Form.StartPosition = "Manual" #allow to set start position manually
# Set start position to top left corner of the main screen
$form.Location.X = -100 
$form.Location.Y = -100
$Form.Text = "QuikNav" # Text value of the form

#create window form
$errorForm = New-Object system.Windows.Forms.Form
#set form size 
$errorForm.Width = 300 # width of the form
$errorForm.Height = 200 # height of the form
$errorForm.MaximizeBox = $false #disable maximize window form
$errorForm.StartPosition = "CenterScreen" #allow to set start position manually
$errorForm.Text = "Error" # Text value of the form

#Add clipboard item to the form
$errorMessage = New-Object System.Windows.Forms.Label 
$errorMessage.Location = New-Object System.Drawing.Size(10,20) 
$errorMessage.Size = New-Object System.Drawing.Size(280,150) 
$errorMessage.AutoSize = $false 
$Font = New-Object System.Drawing.Font("Arial",11)
$errorForm.Font = $Font 
$errorForm.Controls.Add($errorMessage)

#Buttons A1 - A6
###############
$buttonA1 = New-Object System.Windows.Forms.Button # create new button
$buttonA1.Location = New-Object System.Drawing.Size(0,0) # position of the button
$buttonA1.Size = New-Object System.Drawing.Size(100,30) # size of the button
$buttonA1.Text = "Model.docx" # Text value on the button
$buttonA1.BackColor = "LightBlue" # Set color for the button
$buttonA1.AutoSize = $true # Adjust button size to text size
$Form.Controls.Add($buttonA1) # Add button to the form

#Add button to the form
$buttonA2 = New-Object System.Windows.Forms.Button 
$buttonA2.Location = New-Object System.Drawing.Size(0,40) 
$buttonA2.Size = New-Object System.Drawing.Size(100,30) 
$buttonA2.Text = "Modelset.docx"
$buttonA2.BackColor = "LightBlue" 
$buttonA2.AutoSize = $true 
$Form.Controls.Add($buttonA2)

#Add button to the form
$buttonA3 = New-Object System.Windows.Forms.Button 
$buttonA3.Location = New-Object System.Drawing.Size(0,80) 
$buttonA3.Size = New-Object System.Drawing.Size(100,30) 
$buttonA3.Text = "THH_Models" 
$buttonA3.BackColor = "LightBlue" 
$buttonA3.AutoSize = $true 
$Form.Controls.Add($buttonA3)

#Add button to the form
$buttonA4 = New-Object System.Windows.Forms.Button 
$buttonA4.Location = New-Object System.Drawing.Size(0,120) 
$buttonA4.Size = New-Object System.Drawing.Size(100,30) 
$buttonA4.Text = "THH_Modelsets"
$buttonA4.BackColor = "LightBlue" 
$buttonA4.AutoSize = $true 
$Form.Controls.Add($buttonA4)

#Add button to the form
$buttonA5 = New-Object System.Windows.Forms.Button 
$buttonA5.Location = New-Object System.Drawing.Size(0,160) 
$buttonA5.Size = New-Object System.Drawing.Size(100,30) 
$buttonA5.Text = "DOCS DEV_Sup" 
$buttonA5.BackColor = "gold" 
$buttonA5.AutoSize = $true 
$Form.Controls.Add($buttonA5)

#Add button to the form
$buttonA6 = New-Object System.Windows.Forms.Button 
$buttonA6.Location = New-Object System.Drawing.Size(0,200) 
$buttonA6.Size = New-Object System.Drawing.Size(100,30) 
$buttonA6.Text = "Archive" 
$buttonA6.BackColor = "gold" 
$buttonA6.AutoSize = $true 
$Form.Controls.Add($buttonA6)

#Button B1 - B6
###############
$buttonB1 = New-Object System.Windows.Forms.Button 
$buttonB1.Location = New-Object System.Drawing.Size(110, 0) 
$buttonB1.Size = New-Object System.Drawing.Size(100,30) 
$buttonB1.Text = "CR Templates" 
$buttonB1.BackColor = "gold" 
$buttonB1.AutoSize = $true 
$Form.Controls.Add($buttonB1)

#Add button to the form
$buttonB2 = New-Object System.Windows.Forms.Button 
$buttonB2.Location = New-Object System.Drawing.Size(110,40) 
$buttonB2.Size = New-Object System.Drawing.Size(100,30) 
$buttonB2.Text = "BB Templates" 
$buttonB2.BackColor = "gold" 
$buttonB2.AutoSize = $true 
$Form.Controls.Add($buttonB2)

#Add button to the form
$buttonB3 = New-Object System.Windows.Forms.Button 
$buttonB3.Location = New-Object System.Drawing.Size(110,80) 
$buttonB3.Size = New-Object System.Drawing.Size(100,30) 
$buttonB3.Text = "AMDR DEV" 
$buttonB3.BackColor = "gold" 
$buttonB3.AutoSize = $true 
$Form.Controls.Add($buttonB3)

$buttonB4 = New-Object System.Windows.Forms.Button 
$buttonB4.Location = New-Object System.Drawing.Size(110,120) 
$buttonB4.Size = New-Object System.Drawing.Size(100,30) 
$buttonB4.Text = "H: Drive" 
$buttonB4.BackColor = "salmon" 
$buttonB4.AutoSize = $true 
$Form.Controls.Add($buttonB4)

$buttonB5 = New-Object System.Windows.Forms.Button 
$buttonB5.Location = New-Object System.Drawing.Size(110,160) 
$buttonB5.Size = New-Object System.Drawing.Size(100,30) 
$buttonB5.Text = "Feedfiles" 
$buttonB5.BackColor = "gold" 
$buttonB5.AutoSize = $true 
$Form.Controls.Add($buttonB5)

$buttonB6 = New-Object System.Windows.Forms.Button 
$buttonB6.Location = New-Object System.Drawing.Size(110,200) 
$buttonB6.Size = New-Object System.Drawing.Size(100,30) 
$buttonB6.Text = "PSscripts" 
$buttonB6.BackColor = "gold" 
$buttonB6.AutoSize = $true 
$Form.Controls.Add($buttonB6)

#Button C1 - C6
###############
$buttonC1 = New-Object System.Windows.Forms.Button 
$buttonC1.Location = New-Object System.Drawing.Size(220,0) 
$buttonC1.Size = New-Object System.Drawing.Size(100,30) 
$buttonC1.Text = "Issue Log.xslm" 
$buttonC1.BackColor = "lightGreen" 
$buttonC1.AutoSize = $true 
$Form.Controls.Add($buttonC1)

$buttonC2 = New-Object System.Windows.Forms.Button 
$buttonC2.Location = New-Object System.Drawing.Size(220,40) 
$buttonC2.Size = New-Object System.Drawing.Size(100,30) 
$buttonC2.Text = "CRModels Reset" 
$buttonC2.BackColor = "Red" 
$buttonC2.AutoSize = $true 
$Form.Controls.Add($buttonC2)

$buttonC3 = New-Object System.Windows.Forms.Button 
$buttonC3.Location = New-Object System.Drawing.Size(220,80) 
$buttonC3.Size = New-Object System.Drawing.Size(100,30) 
$buttonC3.Text = "PROD" 
$buttonC3.BackColor = "gold" 
$buttonC3.AutoSize = $true 
$Form.Controls.Add($buttonC3)

$buttonC4 = New-Object System.Windows.Forms.Button 
$buttonC4.Location = New-Object System.Drawing.Size(220,120) 
$buttonC4.Size = New-Object System.Drawing.Size(100,30) 
$buttonC4.Text = "Tor_Std_RptS" 
$buttonC4.BackColor = "lightBlue" 
$buttonC4.AutoSize = $true 
$Form.Controls.Add($buttonC4)

$buttonC5 = New-Object System.Windows.Forms.Button 
$buttonC5.Location = New-Object System.Drawing.Size(220,160) 
$buttonC5.Size = New-Object System.Drawing.Size(100,30) 
$buttonC5.Text = "Ibutton12" 
$buttonC5.BackColor = "lightGray" 
$buttonC5.AutoSize = $true 
$Form.Controls.Add($buttonC5)

$buttonC6 = New-Object System.Windows.Forms.Button 
$buttonC6.Location = New-Object System.Drawing.Size(220,200) 
$buttonC6.Size = New-Object System.Drawing.Size(100,30) 
$buttonC6.Text = "Ibutton12" 
$buttonC6.BackColor = "lightGray" 
$buttonC6.AutoSize = $true 
$Form.Controls.Add($buttonC6)

#ALIAS FUNCTIONS
#################Implement the event handlers here#############
function Open-Model  {
     try {
         Invoke-Item -ErrorAction 'Stop' -Path "fgggggggg" 
     } catch [System.Exception] {
        $errorMessage.Text = "$_"
        $errorForm.ShowDialog()
     }
}

#Event Handler Button on-click
$buttonA1.Add_Click({Open-Model})
$buttonA2.Add_Click({Open-Modelset})
$buttonA3.Add_Click({Open-THH-Models})
$buttonA4.Add_Click({Open-THH-Modelsets})
$buttonA5.Add_Click({Open-DEV})
$buttonA6.Add_Click({Open-Archive})

$buttonB1.Add_Click({Open-CRTemplates})
$buttonB2.Add_Click({Open-BBTemplates})
$buttonB3.Add_Click({Open-AMDR-DEV})
$buttonB4.Add_Click({Open-H})
$buttonB5.Add_Click({Open-Feedfiles})
$buttonB6.Add_Click({Open-PSscripts})

$buttonC1.Add_Click({Open-IssueLog})
$buttonC2.Add_Click({Reset-Models})
$buttonC3.Add_Click({Open-PROD})
$buttonC4.Add_Click({Open-Models-Toronto})
$buttonC5.Add_Click({})
$buttonC6.Add_Click({})

#Open the form - this line has to be the last
$Form.ShowDialog()
