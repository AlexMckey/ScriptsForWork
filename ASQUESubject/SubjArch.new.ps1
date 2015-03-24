param(  [string] $ImpPath = "\\as2-volgd\ingis\ImpExp",
        [string] $ArchPath = "\\as2-volgd\ingis\Archive\Subject",
		[string] $ScriptPath = "\\as2-volgd\ingis\EXE\Script",
		[bool]   $Log = $true)

# метка субъекта - макет 800X0
function GetLabel800X0([IO.FileInfo] $fi)
{
	return ($fi.Name.Split("_")[2]).Substring(0,6)
}

# метка субъекта - макет 51070
function GetLabel51070([IO.FileInfo] $fi)
{
	return $fi.LastWriteTime.AddMonths(-1).ToString("yyyyMM")
}

# метка субъекта - макет АСКП
function GetLabelASKP([IO.FileInfo] $fi)
{
	$fy = $fi.LastWriteTime.ToString("yyyy")
	$fm = $fi.BaseName.Substring($fi.BaseName.Length-4,2)
	return $fy+$fm
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
		if ($Log) {"`tYear: " + $lbl.Substring(0,4) + "`t Month: " + $lbl.Substring(4) + "`t ArchFile: " + $archDir + "\$lbl.zipr  - Complite" >> $fl}
		"`tYear: " + $lbl.Substring(0,4) + "`t Month: " + $lbl.Substring(4) + "`t ArchFile: " + $archDir + "\$lbl.zip  - Complite"
	}
}

# обработка директории каждого субъекта
function DoSubject80020([IO.DirectoryInfo] $SubjDir)
{
	# обработка всех файлов формата XML - 80020
	$dn = $SubjDir.Name
	if ($Log) {"`t`tСубъект: " + $dn >> $fl}
	"`t`tСубъект: " + $dn
	$files = $SubjDir.GetFiles("*80020*.*") | ? { $_.Name -notlike "*$tds*"}
	if (($files -ne $null) -and ($files.Count -ne 0)) {ArchFiles $files $dn "GetLabel800X0"}
	# обработка всех файлов формата XML - 80030
	$files = $SubjDir.GetFiles("*80030*.*",[IO.SearchOption]::AllDirectories) | ? { $_.Name -notlike "*$tds*"}
	if (($files -ne $null) -and ($files.Count -ne 0)) {ArchFiles $files "$dn\80030" "GetLabel800X0"}
	# обработка всех файлов формата XML - 80040
	$files = $SubjDir.GetFiles("*80040*.*",[IO.SearchOption]::AllDirectories) | ? { $_.Name -notlike "*$tds*"}
	if (($files -ne $null) -and ($files.Count -ne 0)) {ArchFiles $files "$dn\80040" "GetLabel800X0"}
}

#$ImpDir = MakeDir($ImpPath)
Set-Alias RAR "$ScriptPath\7z.exe"
$td = [DateTime]::Today
$tds = $td.ToString("yyyyMM")
$tdp = $td.AddMonths(-1).ToString("yyyyMM")
$ddmmyy = $td.ToString("dd.MM.yyyy")
$yymm = $td.ToString("yyyy.MM")
$yy = $ddmmyy.Substring(6)
# проверка существования файла с логами работы скрипта, если его нету - создадим
if ($Log) {
	$fl = "$ImpPath\SubjArch.$yy.log"
	if (-not (Test-Path $fl)) {"Лог работы скрипта архивирования файлов субъектов ОРЭ `t-`t $yy`n`t`t(макеты 80020, 80030, 80040, 51070, АСКП)`n" > $fl}
}
"Лог работы скрипта архивирования файлов субъектов ОРЭ `t-`t $yy`n`t`t(макеты 80020, 80030, 80040, 51070, АСКП)`n"

#  очистка директории для удаления
$delDir = Join-Path $ImpPath "BeDeleted"
Get-ChildItem $delDir -Recurse | ? {$_.LastWriteTime -lt $td.AddDays(-$td.Day+1).AddMonths(-1)} | Remove-Item
if ($Log) {"`t`tДиректория для удаления: $delDir `t- почищена" >> $fl}
"`t`tДиректория для удаления: $delDir `t- почищена"

# обработка всех директория с данными субъектов
if ($Log) {"`t$ddmmyy - Архивирование данных субъектов `t Директория назначения: $ArchPath" >> $fl}
"`t$ddmmyy - Архивирование данных субъектов `t Директория назначения: $ArchPath"
$dirs = Get-ChildItem $ImpPath -Exclude "Import*","51070","BeDeleted","ASKP" | ? {$_.PSIsContainer}
$dirs | % { DoSubject80020($_) }
if ($Log) {"`t---- Данные субъектов обработаны" >> $fl}
"`t---- Данные субъектов обработаны"

# обработка макетов 51070
if ($Log) {"`t$ddmmyy - Макет 51070 `t Директория назначения: $ArchPath\51070" >> $fl}
"`t$ddmmyy - Макет 51070 `t Директория назначения: $ArchPath\51070"
$files = Get-ChildItem "$ImpPath\51070" -Include "*.xml*" -Recurse | ? {$_.LastWriteTime -lt $td.AddDays(-$td.Day+1)}
if (($files -ne $null) -and ($files.Count -ne 0)) {ArchFiles $files "51070" "GetLabel51070"}
# удалить лишние поддериктории и файлы
Get-ChildItem "$ImpPath\51070" -Exclude "*.xml" -Recurse | ? {!$_.PSIsContainer} | Move-Item "$ImpPath\BeDeleted\51070" -Force
Get-ChildItem "$ImpPath\51070" -Recurse | ? {$_.PSIsContainer} | Remove-Item -Force
if ($Log) {"`t---- Директория макета 51070 обработана" >> $fl}
"`t---- Директория макета 51070 обработана"

# обработка макетов АСКП
if ($Log) {"`t$ddmmyy - Макет АСКП `t Директория назначения: $ArchPath\ASKP" >> $fl}
"`t$ddmmyy - Макет АСКП `t Директория назначения: $ArchPath\ASKP"
$files = Get-ChildItem "$ImpPath\ASKP" -Recurse | ? {!$_.PSIsContainer -and ($_.LastWriteTime -lt $td.AddDays(-$td.Day+1))}
if (($files -ne $null) -and ($files.Count -ne 0)) {ArchFiles $files "ASKP" "GetLabelASKP"}
if ($Log) {"`t---- Директория макета АСКП обработана" >> $fl}
"`t---- Директория макета АСКП обработана"

# обработка других файлов в директории
if ($Log) {"`t$ddmmyy - Логи и конфиги `t Директории назначения: $ArchPath\LOG  и $ArchPath\CFG" >> $fl}
"`t$ddmmyy - Логи и конфиги `t Директории назначения: $ArchPath\LOG  и $ArchPath\CFG"
$files = Get-ChildItem $ImpPath -Exclude "SubjCopyTbl*.xls*","*$yymm.log" | ? {!$_.PSIsContainer}
#  обработка мусора
$ffd = $files | ? {$_.Extension -ne ".log"}
$ffd | Move-Item -Destination $delDir
if ($Log -and ($ffd -ne $null) -and ($files.Count -ne 0)) {
	"``ttФайлы:" >> $fl
	$ffd | % {"`t`t" + $_.Name >> $fl}
	"`t`tперемещены в директорию на удаление: $delDir" >> $fl
}
if (($ffd -ne $null) -and ($files.Count -ne 0)) {
	"`t`tФайлы:"
	$ffd | % {"`t`t" + $_.Name}
	"`t`tперемещены в директорию на удаление: $delDir"
}
#  обработка логов
$files = Get-ChildItem $ImpPath -Exclude "SubjCopyTbl*.xls*","*$yymm.log" | ? { -not $_.PSIsContainer}
if (($files -ne $null) -and ($files.Count -ne 0)) {
	RAR u -tzip -bd -y "$ArchPath\LOG\log.$yy.rar" $files >> $null
	$files | Remove-Item -Force
}
if ($Log) {"`t`tФайлы логов перемещены в архив: $ArchPath\LOG\log.$yy.rar" >> $fl}
"`t`tФайлы логов перемещены в архив: $ArchPath\LOG\log.$yy.rar"
#  копируем файл с настройками
# Copy-Item -Path "$ImpPath\SubjCopyTbl*.xls*" -Destination (Join-Path $ArchPath "CFG\SubjCopyTbl.xlsx") -Force
RAR u -tzip -bd -y "$ArchPath\CFG\SubjCopyTbl-$ddmmyy.zip" "$ImpPath\SubjCopyTbl.xlsx" >> $null
if ($Log) {"`t---- Логи и конфиги обработаны" >> $fl}
"`t---- Логи и конфиги обработаны"
if ($Log) {"-----------------------------------------------------------------------------------------------------------------------------------------------" >> $fl}
"-----------------------------------------------------------------------------------------------------------------------------------------------"