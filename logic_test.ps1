#Logic testing

$InceptionDateLess10 = $TRUE
$BoardAccount = $FALSE
$SI = $FALSE

If($SI = $False) {
    If($InceptionDateLess10 -and $BoardAccount) {
        write-host "No Show"
        }
    Else {
        Write-host "Show"
        }
}
Else {
      If($InceptionDateLess10) {
        write-host "No Show"
        }
      Else {
        Write-host "Show"
        }
}
