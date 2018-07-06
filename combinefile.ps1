# CR_Feedfile.ps1
# Version 1.0
# Last modified on: May 25, 2018

add-type -assemblyname system.windows.forms
add-type -assemblyname system.drawing

$form = new-object system.windows.forms.form
$form.text = 'data entry form'
$form.size = new-object system.drawing.size(300,200) 
$form.startposition = 'centerscreen'

$okbutton = new-object system.windows.forms.button
$okbutton.location = new-object system.drawing.point(75,120)
$okbutton.size = new-object system.drawing.size(75,23)
$okbutton.text = 'ok'
$okbutton.dialogresult = [system.windows.forms.dialogresult]::ok
$form.acceptbutton = $okbutton
$form.controls.add($okbutton)

$cancelbutton = new-object system.windows.forms.button
$cancelbutton.location = new-object system.drawing.point(150,120)
$cancelbutton.size = new-object system.drawing.size(75,23)
$cancelbutton.text = 'cancel'
$cancelbutton.dialogresult = [system.windows.forms.dialogresult]::cancel
$form.cancelbutton = $cancelbutton
$form.controls.add($cancelbutton)

$label = new-object system.windows.forms.label
$label.location = new-object system.drawing.point(10,20)
$label.size = new-object system.drawing.size(280,20)
$label.text = "please enter short_name_bk (comma-seperated):"
$form.controls.add($label)

$textbox = new-object system.windows.forms.textbox
$textbox.location = new-object system.drawing.point(10,40)
$textbox.size = new-object system.drawing.size(260,20)
$form.controls.add($textbox)

#Error Message
$errorMessage = New-Object System.Windows.Forms.Label 
$errorMessage.Location = New-Object System.Drawing.Size(10,70) 
$errorMessage.Size = New-Object System.Drawing.Size(270,40) 
$errorMessage.AutoSize = $false 

$form.topmost = $true

$form.add_shown({$textbox.select()})
$result = $form.showdialog()

# Event handler: when user click-on 'ok'
if ($result -eq  [system.windows.forms.dialogresult]::ok) {

    # store user input in x for validation
    $x = $textbox.text

    # input must contains only letters, numbers, commas or spaces and less than 200 in length
    # if input is invalid, user can choose to try again or stop the script
    while($x -match "[^a-zA-Z0-9, -]" -or $x.Length -gt 200) {
         
         # display message when input is invalid
         $form.add_shown({$textbox.select()})
         $errorMessage.Text = "Error - Your input contains invalid characters or exceeds 200 characters" 
         $errorMessage.ForeColor = "Red"
         $form.controls.add($errorMessage)
         
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

    # store name of the file to be read from AMDR 1-year Archive
    # file is sorted by last modified time and greater size
    
    $crportfile = [string]
    $crportfile = @(get-childitem "ili" -filter *_cr_portia_portfolio*.txt |
         sort -property lastwritetime, length -descending |
         select -first 1 -expandproperty name)

    $inputfile = "hgh"
    $inputfile = "H:\Pages\feed files\cr_unity_results.txt"

   # write-host "Read in from $crportfile" 


    <################ Change output file path here ####################>

    $outputfile = "H:\Pages\Feed Files\cr_unity_inv_results.txt"
    
    <##################################################################>
    
    <# read the header row and store into output file. This will replace
    the file if one already exists #>
    get-content $inputfile | select -first 1 | out-file -filepath $outputfile

    <# for each item in user's input, find matches row in the archive file  
    and append to the newly created feed file #>
    foreach ($item in $portcodes) {
        get-content $inputfile | where-object{($_ -match $item)} |
            out-file -append -Encoding ascii -filepath $outputfile |
            select -First 1 | out-file -append -filepath $outputfile
    }

    # open the file, this line can be commented out 
    invoke-item -path $outputfile
}


# Credits:
# https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/creating-a-custom-input-box?view=powershell-6
