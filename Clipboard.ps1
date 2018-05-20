<#
Script Name: clipboard.ps1
Last Modified: May 14, 2018
Verions: 1.1
Description: The script inputs csv file and parse each item into clipboard
Specifications:
Input csv file must follow format, please check the sample input file
Features to work on:
Enable previous button at end of file
#>



$filename = $Args[0] #set filename to user input parameter
$fileExt=[System.IO.Path]::GetExtension($filename) #get the input file type

# Check for the correct number of parameters
if ($Args.Count -ne 1) {
	write-host -foregroundcolor yellow "Missing file name. Please try again"
	exit 1
}
elseif ($Args.Count -gt 1) {
	write-host -foregroundcolor yellow  "Can only read in one file at a time. Please try again";
	exit 1
}
# Check for correct file type
elseif ($fileExt -ne ".csv") {
	write-host -foregroundcolor yellow "File is not in csv (Comma Delimited) format. Please try again";
	exit 1
}

# Read in from csv with header Index and Value fields
try {
	$items = import-csv $filename -header Index, Value
}
catch {
	write-host -foregroundcolor yellow "Cannot open file. Please check the file name. There should be no .\ before the filename"
	exit 2
}

$global:i = 0 #count variale

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#create window form
$Form = New-Object system.Windows.Forms.Form
$Form.Width = 350
$Form.Height = 200 
$form.MaximizeBox = $false 
$Form.StartPosition = "CenterScreen" 
$Form.Text = "Current Clipboard Item"
$Form.BackColor = "Black"

#Add clipboard item to the form
$clipboardItem = New-Object System.Windows.Forms.Label 
$clipboardItem.Location = New-Object System.Drawing.Size(10,20) 
$clipboardItem.Text = "Reading from: $filename"
$clipboardItem.AutoSize = $true 
$clipboardItem.ForeColor = "Aquamarine"
$Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
$Form.Font = $Font 
$Form.Controls.Add($clipboardItem)

#Add button to the form
$button = New-Object System.Windows.Forms.Button 
$button.Location = New-Object System.Drawing.Size(110,120) 
$button.Size = New-Object System.Drawing.Size(100,30) 
$button.Text = "Start parsing" 
$button.AutoSize = $true 
$button.BackColor = "cornsilk"
$Form.Controls.Add($button)

#Add button to the form
$nextbtn = New-Object System.Windows.Forms.Button 
$nextbtn.Location = New-Object System.Drawing.Size(175,120) 
$nextbtn.Size = New-Object System.Drawing.Size(80,30) 
$nextbtn.Text = "Next" 
$nextbtn.AutoSize = $true 
$nextbtn.BackColor = "cornsilk"

#Add button to the form
$prevbtn = New-Object System.Windows.Forms.Button 
$prevbtn.Location = New-Object System.Drawing.Size(85,120) 
$prevbtn.Size = New-Object System.Drawing.Size(80,30) 
$prevbtn.Text = "Back" 
$prevbtn.AutoSize = $true 
$prevbtn.BackColor = "cornsilk"


function Next-to-clipboard { 

    if ($global:i -lt $items.count) {

        if ($global:i -eq 0) {
            $Form.Controls.Remove($button)
            $Form.Controls.Add($nextbtn)
            $Form.Controls.Add($prevbtn)
        }

        #set text to clipboard
        $items | where-object {$_.Index -eq $global:i} | select -Unique -ExpandProperty Value | clip

        #change the label to current clipboard text
        $clipboardItem.Text = [Windows.Forms.Clipboard]::GetText();
        $global:i++ # increment i for next reading
    }
    else {
        $clipboardItem.Text = "This is the end of file. Bye ヾ( ＾ - ＾ )ノ"
        $clipboardItem.ForeColor = "Yellow"
        $Form.Controls.Remove($prevbtn)
        $Form.Controls.Remove($nextbtn)
        $button.Text = "Close" 
        $form.CancelButton = $button
        $button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($button)
    }
} 

function Previous-to-Clipboard { 
    
    $global:i -= 2 # increment i for next reading

    while ($global:i -lt 0) {
        $global:i = 0

        #set text to clipboard
        $items | where-object {$_.Index -eq 0} | select -Unique -ExpandProperty Value | clip

        #change the label to current clipboard text
        $clipboardItem.Text = [Windows.Forms.Clipboard]::GetText()
    }

    #set text to clipboard
    $items | where-object {$_.Index -eq $global:i} | select -Unique -ExpandProperty Value | clip

    #change the label to current clipboard text
    $clipboardItem.Text = [Windows.Forms.Clipboard]::GetText()
    $global:i++      
}


#Event Handler Button on-click
$button.Add_Click({Next-to-clipboard})
$nextbtn.Add_Click({Next-to-clipboard})
$prevbtn.Add_Click({Previous-to-Clipboard})

#Open the form
$Form.ShowDialog()


#Credits:
#https://gallery.technet.microsoft.com/scriptcenter/How-to-build-a-form-in-7e343ba3
#http://www.heikniemi.net/hardcoded/2010/01/powershell-basics-1-reading-and-parsing-csv
