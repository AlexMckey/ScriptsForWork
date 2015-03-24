param(  [int] $AfterDays = 3,
		[string] $WavPath = "\\10.87.0.85\Records$",
		[string] $MP3Path = "\\intra-volgd\ControlPeregovor")
function CopyMp3Files(
		[string] $Disp = (Read-Host -Prompt "Введите диспетчера (ФамилияИО): "),
        [DateTime] $Date = ([DateTime]::Today.AddDays(-1)),
		[int] $Smena = 1)
{		
	$Channel = if ($StDisp -notcontains $Disp) {2} else {1}
	$d1 = $Date.ToString("yyyy_MM_dd")
	$d2 = if ($Smena -eq 1) {$d1} else {$Date.AddDays(1).ToString("yyyy_MM_dd")}
	$d3 = if ($Smena -eq 1) {$Date.AddHours(8)} else {$Date.AddHours(20)}
	$d4 = if ($Smena -eq 1) {$Date.AddHours(20)} else {$Date.AddDays(1).AddHours(8)}
	$dir1 = "$WavPath\$Channel"
	$dir2 = "$MP3Path\$Disp\$d1-$Smena"
	if (-not (Test-Path $dir2)) { md $dir2 -Force | Out-Null}
	$f = dir "$dir1\*" -Recurse -Include "$d1`_*.wav","$d2`_*.wav"
	$f |
	? { ($_.LastWriteTime -ge $d3) -and ($_.LastWriteTime -lt $d4) } |
	% { $_.CopyTo("$dir2\" + $_.Name.Replace($_.Extension,".mp3")) }
	[DateTime]::Now.ToString("dd.MM.yyyy HH:mm") + `
	"`t - Записи переговоров $Disp за $d1 ($Smena смена) - скопированы" + "`n" `
	>> "$MP3Path\AutoCopyScript.log"
}

$WavDisk = "V:"
$NetD = New-Object -com WScript.Network
if (Test-Path $WavDisk) { $NetD.RemoveNetworkDrive($WavDisk,"1","1") }
$NetD.MapNetworkDrive($WavDisk,$WavPath,$false,"1","1")
$StDisp = "АнаненкоАН","СоколовАВ","ТельмановНМ","ФайзулинМФ","ПокручинАА","МихальковДВ"
$data = Import-Clixml "$MP3Path\ControlTable.xml"
$data | 
#Расскоментировать следующую строчку и закоментировать после нее для ручного запуска скрипта,
#для копирования всех записей с начала месяца, до текущей даты
#? {$_.Дата -le [DateTime]::Today.AddDays(-$AfterDays+1-$_.Смена)} |
? {$_.Дата -eq [DateTime]::Today.AddDays(-$AfterDays+1-$_.Смена)} |
% {CopyMp3Files $_.Диспетчер $_.Дата $_.Смена}
$NetD.RemoveNetworkDrive($WavDisk,"1","1")
