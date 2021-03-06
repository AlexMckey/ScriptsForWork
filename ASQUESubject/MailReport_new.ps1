param(
    [string]$SubLbl = "Unknown",
    [string[]]$MailFileTo = "mav@volgograd.so-ups.ru",
    [string[]]$ReportTo = "mav@volgograd.so-ups.ru",
    [string]$RepDir = "h:\Отдел сопровождения рынка\Отчеты\Для согласования",
    [string]$ScriptDir = "J:\EXE\Script"
)

# *** Проверяем наличие файла lastsend.xml и, в случае отсутствия, создаем его ***
$LastSendFilePath = "$RepDir\$SubLbl.lastsend"
If (-not (Test-Path $LastSendFilePath)){(Get-Date).Date | Export-Clixml $LastSendFilePath}

[System.IO.FileInfo[]] $FilesToMail = @()
$FilesToMail = Get-ChildItem "$RepDir\*$SubLbl*.xls" | where {$_.lastwritetime -ge (Import-Clixml $LastSendFilePath)}

if ((?? {$FilesToMail.Length} {0}) -eq 0)
{
  "ScriptRes:Ok:NoNeed"
  exit
}

$Body = "Данные для согласования за: `n"+$($FilesToMail | % {$_.Name.Substring(5,6)} | Sort)
$Cnt = $FilesToMail.Length

$FilesZipped = $false
if ($Cnt -gt 5)
{
  $FilesToMail | Write-Zip -OutputPath "$RepDir\$SubLbl.zip" -FlattenPaths -Level 9 
  $FilesToMail = "$RepDir\$SubLbl.zip"
  $FilesZipped = $true
}

# Mail -To $MailFileTo -Subject "Согласование генерации" -Body $Body -AttachmentPath $FilesToMail
& "$ScriptDir\Send-SMTPMail.ps1" -To $MailFileTo -Subject "Согласование генерации" -Body $Body -Attachment $FilesToMail -Server 'ex1-volgd' -From 'asqe@volgograd.so-ups.ru'

$Body = $Body+"`nНаправлены субъекту $SubLbl на согласование..."

# Mail -To $ReportTo -Subject "Согласование генерации" -Body $Body
& "$ScriptDir\Send-SMTPMail.ps1" -To $ReportTo -Subject "Согласование генерации" -Body $Body -Server 'ex1-volgd' -From 'asqe@volgograd.so-ups.ru'

if ($Cnt -ne 0)
{
  "ScriptRes:Ok:$Cnt"
}

if ($FilesZipped) {Del "$RepDir\$SubLbl.zip"}
Get-Date | Export-Clixml $LastSendFilePath