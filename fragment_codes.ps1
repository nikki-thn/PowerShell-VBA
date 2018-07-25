##############################################################################

#Calculate the month-end date, quarter-end date and such

$PriorMonth = (Get-Date).AddMonths(-1)
$CurrentDay = [System.DateTime]::DaysInMonth($PriorMonth.Year, $PriorMonth.Month)
$CurrentMonth = (Get-Date).AddMonths(-1).Month
$CurrentYear = (Get-Date).AddMonths(-1).Year 

$Today = (Get-Date).AddMonths(-2)
$Quarter = [Math]::Ceiling($Today.Month / 3)
$StartDate = ($Quarter * 3 - 2).ToString() + "/1/" + $Today.Year.ToString()
$LastDay = [DateTime]::DaysInMonth([Int]$Today.Year.ToString(),[Int]($Quarter * 3))
$EndDate = ($Quarter * 3).ToString() + "-$LastDay-" + $Today.Year.ToString()

$sqlQuery = "select * from core.time_D where cal_date_bk = '2018-07-14'" 
$dateNum = (Invoke-Sqlcmd -Query $sqlQuery -ServerInstance "####" -Database "###" -Username "####" -Password "####").time_id

write-host $dateNum.toString()
#write-host $Today $Quarter

##############################################################################
#Concat files 
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
    
   #Logic testing

$hf = $TRUE
$hgg = $FALSE
$SI = $FALSE

If($SI = $False) {
    If($fhhf -and $hfh) {
        write-host "No Show"
        }
    Else {
        Write-host "Show"
        }
}
Else {
      If($hfgh) {
        write-host "No Show"
        }
      Else {
        Write-host "Show"
        }
}
  ############################################################################## 
