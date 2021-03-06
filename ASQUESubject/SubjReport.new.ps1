param([string] $RepPath = "\\appsrv-volgd\c$\Program Files\LocalEscortService\OutputFiles",
	  [string] $ScriptPath="\\as2-volgd\ingis\EXE\Script\",
	  [string] $ArchPath = "\\as2-volgd\ingis\Archive\Subject",
	  [string] $ImpPath = "\\as2-volgd\ingis\ImpExp",
      [string] $SubjCopyTblFile = "SubjCopyTbl",
	  [bool]   $Log = $true)

# метка субъекта - макет 800X0
function GetLabel800X0([IO.FileInfo] $fi)
{
	return ($fi.Name.Split("_")[2]).Substring(0,6)
}

# ИНН субъекта - макет 800X0
function GetINN800X0([IO.FileInfo] $fi)
{
	return ($fi.Name.Split("_")[1])
}

function MakeDir([String] $dir)
{
  if (Test-Path $dir -PathType Container)
    {$dir+"\*"}
  else
    {(Split-Path $dir) + "\*"}
}

# архивирование файлов
function ArchFiles([IO.FileInfo[]] $ff, [string] $di, $funk)
{
	$lbls = $ff | % {&$funk($_)} | select -Unique
	$lbls | %{
		$lbl = $_
		$archDir = Join-Path "$ArchPath\$di" $lbl.Substring(0,4)
		if (-not (Test-Path $archDir)) { md $archDir > null}
		$fls = $ff | ? {(&$funk $_) -eq $lbl} | % {$_.FullName}
		RAR u -tzip -bd -y "$archDir\$lbl" $fls >> $null
		$fls | Move-Item "$ImpPath\BeDeleted\$di"
		if ($Log) {"`tYear: " + $lbl.Substring(0,4) + "`t Month: " + $lbl.Substring(4) + "`t ArchFile: " + $archDir + "\$lbl.zip  - Complite" >> $fl}
		"`tYear: " + $lbl.Substring(0,4) + "`t Month: " + $lbl.Substring(4) + "`t ArchFile: " + $archDir + "\$lbl.zip  - Complite"
	}
}

function dmy()
{
	[DateTime]::Today.ToShortDateString() + " " + [DateTime]::Now.ToShortTimeString() + " = "
}

Set-Alias RAR "$ScriptPath\7z.exe"
$td = [DateTime]::Today
$tds = $td.ToString("yyyyMM")
$ddmmyy = $td.ToString("dd.MM.yyyy")
$yymm = $td.ToString("yyyy.MM")
$yy = $ddmmyy.Substring(6)
# проверка существования файла с логами работы скрипта, если его нету - создадим
if ($Log) {
	$fl = "$ImpPath\SubjReport.$yymm.log"
	if (-not (Test-Path $fl)) {"Лог работы скрипта отправки макетов 80021 субъектам ОРЭ`n" > $fl}
}
"Лог работы скрипта отправки макетов 80021 субъектам ОРЭ `t-`t $yy`n"

# инициализация путей фалов
$fcsv = Join-Path -Path $ScriptPath -ChildPath "$SubjCopyTblFile.csv"
$RepDir = MakeDir($RepPath)

# загрузка файла с таблицей копирования субъектов
$RepTbl = Import-Csv $fcsv -Delimiter ";"

# удаляем мусор
Move-Item $RepPath -Destination "$ImpPath\BeDeleted" -Exclude "*.xml"

# обработка макетов 80021
$sendINN = $RepTbl | ? { $_.TOEMAIL -ne $null -and $_.TOEMAIL -ne ""}
$INNlist = $sendINN | % {$_.INN}
#  архивирование макетов, отправка которых не требуется
if ($Log) {"`t$ddmmyy - Макет 80021 - Отправка не требуется `t Архив: $ArchPath\80021\NotSended" >> $fl}
"`t$ddmmyy - Макет 80021 - Отправка не требуется `t Архив: $ArchPath\80021\NotSended"
$files = Get-ChildItem $RepDir -Include "*80021*.xml" | ? { $INNlist -notcontains (GetINN800X0($_))}
if ((?? {$files.Count} {0}) -ne 0) {ArchFiles $files "80021\NotSended" "GetLabel800X0"}
#   отправка макетов субъектам по почте
if ($Log) {"`t$ddmmyy - Макет 80021 - Отправка субъектам `t Архив: $ArchPath\80021" >> $fl}
"`t$ddmmyy - Макет 80021 - Отправка субъектам `t Архив: $ArchPath\80021"
$files = Get-ChildItem $RepDir -Include "*80021*.xml" | ? { $INNlist -contains (GetINN800X0($_))}
if ((?? {$files.Count} {0}) -ne 0) {
	$files | % {
		$INN = GetINN800X0($_)
		$toM = $sendINN | ? {$_.INN -eq $INN} | % {$_.TOEMAIL}
		# временная заглушка - потом закомментировать или удалить
		$toM = "mav@volgograd.so-ups.ru"
		& "$ScriptPath\Send-SMTPMail.ps1" -To $toM -Subject "Maket 80021 - VRDU" -Attachment $_ -Server 'ex2-volgd' -From 'asqe@volgograd.so-ups.ru'
	}
	ArchFiles $files "80021\Sended" "GetLabel800X0"
}
if ($Log) {"`t---- Макеты 80021 обработаны" >> $fl}
"`t---- Макеты 80021 обработаны"


# копируем файлы формата XML800X0
$XML800X0Files = Get-ChildItem $ImpDir -Include "*800?0*.xml"	
if ((($XML800X0Files | Measure-Object).Count) -gt 0)
{
	foreach($file in $XML800X0Files)
	{
		$b = $false
		
		$t = $file.Name -split "_"
		$m = $t[0]
		if ($m.Length -gt 5) {$m = $m.Substring(1)}
		$i = $t[1]
		$d = $t[2]
		$si = $CopyTbl | ? {$_.INN -eq $i}
		# Отельно обработать Алюминьку
		if (($si -ne $null) -and ($si.DIR -eq "ALZ"))
		{
			$alz = Get-Content $file
			$alzs = ($alz | % { $_.Replace('0 "' , '0"') })
			$alzs | Out-File -FilePath $file -Force -Encoding "default"
		}
		if (($si -ne $null) -and ($si.OIK -lt 0))
		{
			Remove-Item $file -Force
			if ($Log) {"$(&dmy)`t$($si.DIR)`t-Removed->`t$($file.Name) !!!!!!!!" >> $fl}
		}
		if (($si -ne $null) -and ($si.OIK -eq 2))
		{
			$u = [int]$si.LTEQGT
			$l = [int]$si.LEN
			$b = if ($u -lt 0) {$file.Length -lt $l} else { if ($u -gt 0) {$file.Length -gt $l} else {$file.Length -eq $l}}
		}
		if (($si -ne $null) -and (($si.OIK -eq 1) -or (($si.OIK -eq 2) -and $b))) 
		{
			Copy-Item $file $OIKPath
			if ($Log) {"$(&dmy)`t$($si.DIR)`t-Copied`t->`t$($file.Name) `tto`tASKUE_OIK_Path" >> $fl}
		}
		$CopyPath = $si.DIR
		if ($m -ne 80020) {$CopyPath += "\" + $m}
		Move-Item $file "$ImpPath\$CopyPath" -Force
		if ($Log) {"$(&dmy)`t$($si.DIR)`t- Moved`t->`t$($file.Name) `tto`t$CopyPath" >> $fl}
	}
	$Cop = $true
}

if ($Log) {"-----------------------------------------------------------------------------------------------------------------------------------------------" >> $fl}
"-----------------------------------------------------------------------------------------------------------------------------------------------"