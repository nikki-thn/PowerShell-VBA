# CR_Feedfile.ps1
# Version 1.0
# Last modified on: May 25, 2018


    $crportfile = get-childitem -path "C:\Users\truonnh\Documents\cr_unity"

    <################ Change output file path here ####################>

    $outputfile = "C:\Users\truonnh\Documents\cr_unity_results.txt"
    
    <##################################################################>

    <# for each item in user's input, find matches row in the archive file  
    and append to the newly created feed file #>
    foreach ($file in $crportfile) {
        get-content -path "C:\Users\truonnh\Documents\cr_unity\$file" | out-file -encoding ascii -append -filepath $outputfile
    }

    # open the file, this line can be commented out 
    invoke-item -path $outputfile



# Credits:
# https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/creating-a-custom-input-box?view=powershell-6
