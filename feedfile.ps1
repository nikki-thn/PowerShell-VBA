# FeedFiles.ps1
# Version 1.3
# Last Modified: May 23rd, 2018

<################ Change output file path here ####################>

$outputfile = "H:\Pages\feed files\cr_portia_portfolio.txt"

# store name of the file to be read from AMDR 1-year Archive
# file is filtered by size (greater than ~2MB) and sorted by last modified time
$crportfile = [string]
$crportfile = @(get-childitem "\\Marchive" -filter *_cr_tfolio.txt |
        where-object {$_.Length -gt 2000000} | sort -property LastWriteTime -descending | 
        select -first 1 -expandproperty name)

 $inputfile = "\\Mlisirchive\$crportfile"
    
<##################################################################>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CR_Portia_Portfolio Feed File'
$form.Size = New-Object System.Drawing.Size(300,300)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,220)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,220)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter portia code(s) below (comma-seperated):'
$form.Controls.Add($label)

#Error Message
$message = New-Object System.Windows.Forms.Label 
$message.Location = New-Object System.Drawing.Size(10,70) 
$message.Size = New-Object System.Drawing.Size(270,120) 
$message.AutoSize = $false 

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true
$form.Add_Shown({$textBox.Select()})

$lastModifiedTime = get-childitem $inputfile | select -ExpandProperty LastWriteTime
$fileSize = [double] 
$fileSize = get-childitem $inputfile | select -ExpandProperty Length
$fileSize = [math]::round(($fileSize/1048576),2)

$message.Text = "Reading in from file`n`nName: $crportfile `n`nLast Modified on: $lastModifiedTime `n`nFile size: $fileSize Mb" 
$form.controls.add($message)

function Handle-OKButton {

    # store user input in x for validation
    $x = $textbox.text

    # input must contains only letters, numbers, commas or spaces and less than 200 in length
    # if input is invalid, user can choose to try again or stop the script
    while($x -match "[^a-zA-Z0-9, ]" -or $x.Length -gt 200) {
         
         # display message when input is invalid
         $form.add_shown({$textbox.select()})
         $message.Text = "Error - Your input contains invalid characters or exceeds 200 characters" 
         $message.ForeColor = "Red"
     
         # retry 
         $result = $form.showdialog() 
         $x = $textbox.text

         # user choose 'cancel' to stop script
         if ($result -eq [system.windows.forms.dialogresult]::cancel) {
            exit 0
         }
    } 

    # process user's input string to components
    $portcodes = @()
    $portcodes = $x.split("{,}").trim() #split string and trim trailing spaces

    #$form.controls.add($message)

    <# read the header row and store into output file. This will replace
    the file if one already existed #>
    get-content $inputfile | select -first 1 | out-file -Encoding ascii -filepath $outputfile

    <# for each item in user's input, find matches row in the archive file  
    and append to the newly created feed file #>
    foreach ($item in $portcodes) {
        $result = @(get-content $inputfile | where-object{($_ -match $item)} |
         select -First 1 | out-file -append -Encoding ascii -filepath $outputfile)
    }

    # open the file
    invoke-item -path $outputfile

    $message.Text = "cr_portia_portfolio.txt has been generated `n`nOutput location: $outputfile"
    $textbox.text = "" 
}

#Event Handler Button on-click
$OKButton.Add_Click({Handle-OKButton})

$form.ShowDialog()
