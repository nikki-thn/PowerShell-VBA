<#
 get-acl '.\Model_Toronto_Std_Rpt setup.docx' | format-list 
 Get-Acl $files | Get-Member -MemberType *Property
 #>
 
#$User =
#$file =

$Acl = Get-Acl $file
$Ar = New-Object  system.security.accesscontrol.filesystemaccessrule($User,"FullControl","Allow")
$Acl.SetAccessRule($Ar)
Set-Acl $file $Acl
