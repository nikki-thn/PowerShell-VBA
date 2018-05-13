 <#
Author: Nikki Truong
Script Name: clipboard.ps1
Last Modified: May 13, 2018
Verions: 1.0
#>

　
$filename = $Args[0]; #set filename to user input parameter

#read-in csv with header Index and i
$items = import-csv $filename -header Index, Value

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
 
