param([string]$BackupType = 'differential',
      [string]$BackupDir = 'D:\backup')
New-Variable ShpAdminToolDir 'c:\Program Files\Common Files\Microsoft Shared\web server extensions\12\BIN\' -Option Constant

$LastSendFilePath = "$BackupDir\Backup.last"
If (-not (Test-Path $LastSendFilePath)){(Get-Date).Date | Export-Clixml $LastSendFilePath}

.\STSADM.EXE -help backup -directory $BackupDir -backupmethod $BackupType -overwrite -quiet

[System.IO.DirectoryInfo[]] $FilesToZip = @()

$FilesToZip = Get-Item "$BackupDir\spbr*" -Exclude "*.xml" | where {$_.lastwritetime -ge (Import-Clixml $LastSendFilePath)}

if ($BackupType -eq 'differential')
{
    Write-Zip $FilesToZip "$BackupDir\SHP-backup.zip" -Append -Quiet
}
else
{
    Write-Zip $FilesToZip "$BackupDir\SHP-backup.zip" -Quiet
}
Remove-Item $FilesToZip -Recurse -Force

Get-Date | Export-Clixml $LastSendFilePath