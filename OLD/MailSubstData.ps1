$Debug = $args[-1] -eq "debug"
Set-Variable RepDir "J:\ImpExp\VE\Замена" -option constant
if ($Debug) {$Line = "-------------------------------------------------------------------------------------------------------------------------------------------"}
if ($Debug) {$DF = "MailReport.debug"}
if ($Debug) {"Debug # Mail-Report #  Date: $([DateTime]::Now)" > $DF}
if ($Debug) {$Line >>$DF}
$a0 = $args[0]
$a1 = $args[1]
$a2 = $args[2]
if ($Debug) {$a0 >>$DF}
if ($Debug) {$a1 >>$DF}
if ($Debug) {$a2 >>$DF}
$SubLbl = ?? {$a0} {"Unknown"}
if ($Debug) {"Label: $SubLbl" >>$DF}
$MailedAttrib = @{}
if ($Debug) {$MailedAttrib >>$DF}
$FilesInDir = Dir "$RepDir\*$SubLbl*.xls" -name
if ($FilesInDir -eq $null)
{
  "ScriptRes:Ok:NoFiles"
  exit
}
if ($Debug) {$Line >>$DF}
if ($Debug) {$FilesInDir >>$DF}
$FilesInDir | % {$MailedAttrib.$_ =$true}
if ($Debug) {$Line >>$DF}
if ($Debug) {Test-Path "$RepDir\$SubLbl.mailed" -PathType leaf >>$DF}
if (Test-Path "$RepDir\$SubLbl.mailed" -PathType leaf)
{
  Type "$RepDir\$SubLbl.mailed" | % {$MailedAttrib.$_=$false}
}
if ($Debug) {$MailedAttrib >>$DF}
if ($Debug) {$Line >>$DF}
$FilesToMail = $MailedAttrib.GetEnumerator() | ? {$_.Value} | % {$_.Key}
if ($Debug) {$FilesToMail >>$DF}
if ($Debug) {$Line >>$DF}
if ($Debug) {($FilesToMail.Length -eq 0) -or ($FilesToMail -eq $null) >>$DF}
if (($FilesToMail.Length -eq 0) -or ($FilesToMail -eq $null))
{
  "ScriptRes:Ok:NoNeed"
  exit
}
#$To = ?: {$args.Length -le 1} {"mav@rduvolgograd.ru"} {$args[1..$args.Length]}
$To = ?? {$a1} {"asqe@rduvolgograd.ru"}
$Body = "Замещение генерации: `n"+$($FilesToMail | % {$idx = $_.IndexOf(" "); "`tпо "+$_.Substring(3,$idx-3)+" за "+$_.Substring($idx+1,10)} | Sort)
$Cnt = ?: {$FilesToMail -is [String]} {1} {$FilesToMail.Length}
if ($Debug) {$To >>$DF; $Body >>$DF; $Cnt >>$DF; $Line >>$DF}
$FilesZipped = $false
$FilesToMail = $FilesToMail | % {"$RepDir\$_"}
if ($Debug) {($FilesToMail.Length -ge 5) -and ($FilesToMail -isnot [String]) >>$DF}
if (($FilesToMail.Length -ge 5) -and ($FilesToMail -isnot [String]))
{
  $FilesToMail | Write-Zip -OutputPath "$RepDir\$SubLbl.zip" -FlattenPaths -Level 9 
  $FilesToMail = "$RepDir\$SubLbl.zip"
  $FilesZipped = $true
}
if ($Debug) {$FilesToMail >>$DF}
if ($Debug) {$Line >>$DF}
Mail -To $To -Subject "Замещение генерации" -Body $Body -AttachmentLiteralPath $FilesToMail
$To = ?? {$a2} {"asqe@rduvolgograd.ru"}
$Body = $Body+"`nНаправлены в ВЭ для замещения данных..."
if ($Debug) {$To >>$DF; $Body >>$DF; $Line >>$DF}
Mail -To $To -Subject "Замещение генерации" -Body $Body
if ($Cnt -ne 0)
{
  $FilesInDir >"$RepDir\$SubLbl.mailed"
  "ScriptRes:Ok:$Cnt"
}
else
{
  "ScriptRes:Bad:$Cnt"
}
if ($FilesZipped) {Del "$RepDir\$SubLbl.zip"}