Set-Variable RepDir "h:\Отдел сопровождения рынка\Отчеты\Для согласования" -option constant
$RepDir >Sasha.debug
Set-Variable Line "------------------------------------------------------------------------------------------------------------------------" -option constant
$SubLbl = $(if ($args[0] -eq $null) {"Unknown"} else {$args[0]})
$SubLbl >>Sasha.debug
$Line >>Sasha.debug
$MailedAttrib = @{}
$FilesInDir = Get-ChildItem "$RepDir\*$SubLbl*.xls" -name
$FilesInDir >>Sasha.debug
$FilesInDir | % {$MailedAttrib.$_=$true}
$Line >>Sasha.debug
Test-Path "$RepDir\$SubLbl.mailed" -PathType leaf >>Sasha.debug
if (Test-Path "$RepDir\$SubLbl.mailed" -PathType leaf)
{
  Get-Content "$RepDir\$SubLbl.mailed" | % {$MailedAttrib.$_=$false}
}
$FilesToMail = $MailedAttrib.GetEnumerator() | ? {$_.Value} | % {$_.Key}
$FilesToMail >>Sasha.debug
$Line >>Sasha.debug
$(($FilesToMail.Length -eq 0) -or ($FilesToMail -eq $null)) >>Sasha.debug
if (($FilesToMail.Length -eq 0) -or ($FilesToMail -eq $null))
{
  "ScriptRes:Ok:NoNeed"
  exit
}
$From =  $(if ($args[1] -eq $null) {"mav@rduvolgograd.ru"} else {$args[1]})
$From =  "mav@rduvolgograd.ru"
$To = $(if ($args.Length -le 2) {"mav@rduvolgograd.ru"} else {$args[2..$args.Length]})
$To = $(if ($args.Length -le 1) {"mav@rduvolgograd.ru"} else {$args[1..$args.Length]})
$Mail = New-Object -COMObject "XSMTP.MailSender"
$Mail.host = "rdulotus"
$Mail.CharSet = "windows-1521"
$Mail.From = $From
$Mail.FromName = "Alex Mc'key"
$To | % {$Mail.AddAddress("$_")}
$Mail.Subject = "Согласование генерации"
$Mail.Body = "Данные для согласования за: `n"+$($FilesToMail | % {$_.Substring(5,6)} | Sort)
$Mail >>Sasha.debug
$FilesZipped = $false
$FilesToMail >>Sasha.debug
$Line >>Sasha.debug
$FilesToMail.Length -ge 5 >>Sasha.debug
$FilesToMail.GetType() >>Sasha.debug
$FilesToMail.GetType().Name >>Sasha.debug
if (($FilesToMail.Length -ge 5) -and ($FilesToMail.GetType().Name -ne "String"))
{
  $FilesToMail | % {"$RepDir\$_"} | Write-Zip -OutputPath "$RepDir\$SubLbl.zip" -FlattenPaths -Level 9 
  $Mail.AddAttachment("$RepDir\$SubLbl.zip")
  $FilesZipped = $true
}
else
{
  $FilesToMail | % {$Mail.AddAttachment("$RepDir\$_")}
}
$FilesZipped >>Sasha.debug
$Cnt = $(if ($FilesToMail.GetType().Name -eq "String") {1} else {$FilesToMail.Length})
if ($Mail.Send() -ne 0)
{
  $FilesInDir >"$RepDir\$SubLbl.mailed"
  "ScriptRes:Ok:$Cnt"
}
else
{
  "ScriptRes:Bad:$Cnt"
}
if ($FilesZipped) {Remove-Item "$RepDir\$SubLbl.zip"}