<#
Author: Nikki Truong
Script Name: clipboard.ps1
Last Modified: May 13, 2018
Verions: 1.0
Description: The script inputs csv file and parse each item into clipboard
Specifications:
Input csv file must follow certain format, please check the sample input file
#>

$filename = $Args[0]; #set filename to user input parameter
$fileExt=[System.IO.Path]::GetExtension($filename); #get the input file type

# Check for the correct number of parameters
if ($Args.Count -ne 1) {
	write-host -foregroundcolor yellow "Missing file name. Please try again";
	exit 1;
}
elseif ($Args.Count -gt 1) {
	write-host -foregroundcolor yellow  "Can only read in one file at a time. Please try again";
	exit 1;
}
# Check for correct file type
elseif ($fileExt -ne ".csv") {
	write-host -foregroundcolor yellow "File is not in csv (Comma Delimited) format. Please try again";
	exit 1;
}

# Read in from csv with header Index and Value fields
try {
	$items = import-csv $filename -header Index, Value
}
catch {
	write-host -foregroundcolor yellow "Cannot open file. Please check the file name"
	exit 2
}

$global:i = 0;

#create window form
$Form = New-Object system.Windows.Forms.Form
$Form.Width = 300
$Form.Height = 200 
$form.MaximizeBox = $false 
$Form.StartPosition = "CenterScreen" 
$Form.Text = "Current Clipboard Item"

#Add clipboard item to the form
$clipboardItem = New-Object System.Windows.Forms.Label 
$clipboardItem.Location = New-Object System.Drawing.Size(10,20) 
$clipboardItem.Text = "File: $filename"
$Font = New-Object System.Drawing.Font("Arial",11) 
$form.Font = $Font 
$Form.Controls.Add($clipboardItem)

#Add button to the form
$Okbutton = New-Object System.Windows.Forms.Button 
$Okbutton.Location = New-Object System.Drawing.Size(95,120) 
$Okbutton.Size = New-Object System.Drawing.Size(100,30) 
$Okbutton.Text = "Start parsing" 
$Form.Controls.Add($Okbutton)

function Copy-to-clipboard { 

    #set text to clipboard
    $items | where-object {$_.Index -eq $global:i} | select -Unique -ExpandProperty Value | clip

    #change the label to current clipboard text
    $clipboardItem.Text = [Windows.Forms.Clipboard]::GetText();

    $Okbutton.Text = "Next" 

    $global:i++;
} 


#Event Handler Button on-click
$Okbutton.Add_Click({Copy-to-clipboard})

#Open the form
$Form.ShowDialog()


#Credits:
#https://gallery.technet.microsoft.com/scriptcenter/How-to-build-a-form-in-7e343ba3
#http://www.heikniemi.net/hardcoded/2010/01/powershell-basics-1-reading-and-parsing-csv
