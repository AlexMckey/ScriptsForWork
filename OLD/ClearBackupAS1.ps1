param(  [string] $BackupPath = "\\as1-volgd\Backup_AIS",
        [int] $BackupCount = 7)
		
$a = Get-ChildItem -Path $BackupPath -Include "*.zip" | Sort-Object LastWriteTime -Descending
$a[$BackupCount..$a.Count] | Remove-Item