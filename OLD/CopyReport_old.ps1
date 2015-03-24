Set-Variable RepDir "h:\Отдел сопровождения рынка\Отчеты\Для согласования" -option constant
#$RepDir >Sasha.debug
#Set-Variable Line "------------------------------------------------------------------------------------------------------------------------" -option constant
#$Line >>Sasha.debug
$SubLbl = $(if ($args[0] -eq $null) {"Unknown"} else {$args[0]})
#$SubLbl >>Sasha.debug
$ToFolder =  $(if ($args[1] -eq $null) {""} else {$args[1]})
#$ToFolder >>Sasha.debug
#$Line >>Sasha.debug
$CopiedAttrib = @{}
$FilesFromDir = Get-ChildItem "$RepDir\*$SubLbl*.xls" -name
#$FilesFromDir >>Sasha.debug
#$Line >>Sasha.debug
$FilesFromDir | % {$CopiedAttrib.$_=$true}
#Test-Path "$ToFolder" -PathType Container >>Sasha.debug
if (Test-Path "$ToFolder" -PathType Container)
{
  $FilesToDir = Get-ChildItem "$ToFolder\*$SubLbl*.xls" -name
  #$FilesToDir >>Sasha.debug
  #$Line >>Sasha.debug
  $FilesToDir | % {$CopiedAttrib.$_=$false}
}
else
{
  "ScriptRes:Bad:NoFolder"
  exit
}
$FilesToCopy = $CopiedAttrib.GetEnumerator() | ? {$_.Value} | % {$_.Key}
#$FilesToCopy >>Sasha.debug
#$Line >>Sasha.debug
#$(($FilesToCopy.Length -eq 0) -or ($FilesToCopy -eq $null)) >>Sasha.debug
if (($FilesToCopy.Length -eq 0) -or ($FilesToCopy -eq $null))
{
  "ScriptRes:Ok:NoNeed"
  exit
}
$From =  "mav@rduvolgograd.ru"
$To = $(if ($args[2] -eq $null) {"mav@rduvolgograd.ru"} else {$args[2]})
$Mail = New-Object -COMObject "XSMTP.MailSender"
$Mail.host = "rdulotus"
$Mail.CharSet = "windows-1521"
$Mail.From = $From
#$Mail.FromName = "Alex Mc'key"
$To | % {$Mail.AddAddress("$_")}
$Mail.Subject = "Согласование генерации"
$Mail.Body = "Выложены новые данные для согласования за: `n"+$($FilesToCopy | % {$_.Substring(13,6)} | Sort)
#$Line >>Sasha.debug
#$Mail >>Sasha.debug
Copy $($FilesToCopy | % {"$RepDir\$_"}) -Destination $ToFolder
if ($Mail.Send() -ne 0)
{
  "ScriptRes:Ok:$($FilesToCopy.Length)"
}
else
{
  "ScriptRes:Unknown:NotSend"
}