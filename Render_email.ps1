$month = (Get-Culture).DateTimeFormat.GetMonthName([datetime]::now.addMonths(+1).Month)

$RecipientsAddr = "Nikki_Truong@manulife.com"
$BodyMsg = "We’re putting in together the reporting schedule for $month, and we have a few questions for your team.<br><br></p>"

#create COM object named Outlook
$Outlook = New-Object -ComObject Outlook.Application
#create Outlook MailItem named Mail using CreateItem() method
$Mail = $Outlook.GetNameSpace("MAPI")
$Mail = $Outlook.CreateItem(0)
$Mail.GetInspector.Activate()
$Signature = $Mail.HTMLBody
[Void]$Mail.Recipients.Add($RecipientsAddr)
#add properties as desired
$Mail.To = "Nikki_Truong@manulife.com"
$Mail.Subject = "New/Closed Accounts - $month"
$Mail.CC = "Nikki_Truong@manulife.com"
$Mail.HTMLBody = "<p style='font:Arial;font-size:11pt'>Hi All, <br><br>" + $BodyMsg + 
    "<span style='font-size:9pt'>This email was automatically generated</span></p>" + $Signature

#$inspector = $mail.GetInspector
#$inspector.Display()
#send message
#$Mail.Send()
#quit and cleanup
#$Outlook.Quit()

#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

#$Mail.save()
#$inspector = $mail.GetInspector

<#

#create COM object named Outlook
$Outlook = New-Object -ComObject Outlook.Application
#create Outlook MailItem named Mail using CreateItem() method
$Mail = $Outlook.GetNameSpace("MAPI")
$Mail = $Outlook.CreateItem(0)
$Mail.GetInspector.Activate()
#add properties as desired
$Mail.To = "nikki_truong@manulife.com"
$Mail.Subject = "subject"
$Signature = $Mail.HTMLBody
$Mail.HTMLBody = "<p>Hello, <br> Testing </p>"
#send message
#$Mail.Send()
#quit and cleanup
#$Outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

$Mail.save()
#$inspector = $mail.GetInspector
#$inspector.Display()
#Credits:
#Found on Spiceworks: https://community.spiceworks.com/how_to/150253-send-mail-from-powershell-using-outlook?utm_source=copy_paste&utm_campaign=growth

#>
