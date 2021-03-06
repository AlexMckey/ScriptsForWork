param([string] $ImpPath = "\\as2-volgd\ingis\ImpExp",
      [string] $OIKPath = "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles",
	  [string] $ScriptPath="\\as2-volgd\ingis\EXE\Script\",
	  [string] $SubjCopyTblFile = "SubjCopyTbl",
	  [string] $ArchDir = "d:\Script",
	  [bool]   $Log = $true)

function MakeDir([String] $dir)
{
  if (Test-Path $dir -PathType Container)
    {$dir+"\*"}
  else
    {(Split-Path $dir) + "\*"}
}

function UnzipFiles([string] $path)
{
	Set-Alias UnZIP "$ArchDir\UnZip.exe"
	foreach($file in $input)
	{
		Write-Host "$path\$file"
		UnZIP -oqq "$path\$file" -d $path
		if ($Log) {"$(&dmy) `t`t- Unzip ->`t$($file.Name)" >> $fl}
		$file.Delete()
	}
}

function UnrarFiles([string] $path)
{
	Set-Alias UnRAR "$ArchDir\UnRar.exe"
	foreach($file in $input)
	{
		Write-Host "$path\$file"
		UnRAR e -ep -o+ -inul -y "$path\$file" $path
		if ($Log) {"$(&dmy) `t`t- Unrar ->`t$($file.Name)" >> $fl}
		$file.Delete()
	}
}

function Un7zFiles([string] $path)
{
	Set-Alias Un7Z "$ArchDir\7Z.exe"
	Push-Location $path
	foreach($file in $input)
	{
		Write-Host "$path\$file"
		Un7Z e -bd -y "$path\$file"
		if ($Log) {"$(&dmy) `t`t-Un7zed ->`t$($file.Name)" >> $fl}
		$file.Delete()
	}
	Pop-Location
}

function dmy()
{
	[DateTime]::Today.ToShortDateString() + " " + [DateTime]::Now.ToShortTimeString() + " = "
}

# если необходимо добавим к директории '\'
$ImpDir = MakeDir("$ImpPath\Import")

# проверка существования файла с логами работы скрипта, если его нету - создадим
$yymm = [DateTime]::Today.ToString("yyyy.MM")
$fl = "$ImpPath\SubjCopy.$yymm.log"
if (-not (Test-Path $fl)) {"Лог работы скрипта копирования файлов субъектов ОРЭ `t-`t $yymm`n     (макеты 80020, 80030, 80040, 51070, АСКП`n" > $fl}

# инициализация путей фалов
$fxml = Join-Path -Path $ScriptPath -ChildPath "ExcelFile.xml"
$fxls = Join-Path -Path $ImpPath -ChildPath "$SubjCopyTblFile.xls*"
$fcsv = Join-Path -Path $ScriptPath -ChildPath "$SubjCopyTblFile.csv"

# в переменную f помещаем объект - excel'евский файл c настройками
$f = Get-ChildItem $fxls
# если такой файл существует то начинаем обрабатывать его
if ($f -ne $null)
{
	# проверяем существует ли объект xml-файл, если да помещаем его в переменную fi
	$fi = if (Test-Path $fxml) {Import-Clixml $fxml}
	$b = Test-Path $fcsv
	if (($fi -eq $null) -or (-not $b) -or ($f.LastWriteTime -ne $fi.LastWriteTime))
	{
		#Открываем в excel-е необходимый нам файл
		$objExcel = New-Object -comobject Excel.Application
		$objWorkbook = $objExcel.Workbooks.Open($f)
		$objExcel.Application.DisplayAlerts = $false
		# функция обработки XLS файла и сохранения данных CSV, 6 - тип CSV
		$xlCSV = 6
		$objWorkbook.SaveAs($fcsv,$xlCSV,$null,$null,$null,$null,$null,"windows-1251") 
		# сохраняем данные в xml-файл
		$f | Select Name, LastWriteTime | Export-Clixml $fxml
		#Закрываем книгу Excel
		$objExcel.Workbooks.Close()
		#Выходим из Excel (вернее даем команду на выход из Excel)
		$objExcel.Application.DisplayAlerts = $true		
		$objExcel.Quit()
		#обнуляем объект
		$objExcel = $null
		#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
		[gc]::collect()
		[gc]::WaitForPendingFinalizers()				
	}
}
	
# загрузка файла с таблицей копирования субъектов
$CopyTbl = Import-Csv $fcsv -Delimiter ";"

# разархивируем архивы
Get-ChildItem "$ImpPath\Import" "*.zip" | UnzipFiles "$ImpPath\Import"
Get-ChildItem "$ImpPath\Import" "*.rar" | UnrarFiles "$ImpPath\Import"
Get-ChildItem "$ImpPath\Import" "*.7z" | Un7zFiles "$ImpPath\Import"

# удаляем мусор
Remove-Item $ImpDir -Include "*.html,*.log,*.jpg,*.xls,*.xlsx".Split(",")

$Cop = $false

# копируем файлы формата XML800X0
$mescfile = Get-ChildItem $ImpDir -Include "*80020*200909.xml" 
$mescfile | % {
if ($_.Length -gt 7Mb) 
{
$_ | Rename-Item -NewName {$_.name -replace '.200909','_2009092'} -Force
}
}
#| ? {$_.Length -gt 7Mb} `
#| % {[xml]$doc  =  Get-Content $_, $doc.ReplaceChild('<inn>2009092</inn>', '<inn>200909</inn>'), $doc.Save($_) `
#| Rename-Item -NewName {$_.name -replace '200909.','2009092.'} -Force }

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
		if (($si -ne $null) -and ($si.DIR -eq "MESC") -and ($file.length -gt 7Mb))
		{
			[xml]$doc  =  Get-Content $file 
			$doc.SelectSingleNode("message/area[inn='200909']/inn").InnerText = "2009092" 
			$doc.Save($file)
			
		#	$mescfile  =  Get-Content $file 
		#	$mesc = ($mescfile | % { $_[14].Replace('<inn>200909</inn>' , '<inn>2009092</inn>') })
		#	$mesc | Out-File -FilePath $file -Force -Encoding "default"
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

$XML51070Files = Get-ChildItem $ImpDir -Include "*.xml" -Exclude "*80020*.xml","*80030*.xml","*80040*.xml"
if ((($XML51070Files | Measure-Object).Count) -gt 0)
{
	foreach($file in $XML51070Files)
	{
		$CopyPath = $ImpPath + "\51070"
		Move-Item $file $CopyPath -Force
		if ($Log) {"$(&dmy)`t51070`t- Moved`t->`t$($file.Name) `tto`t51070" >> $fl}
	}
	$Cop = $true
}

$ASKPFiles = Get-ChildItem  $ImpDir -Include "vtz*","az*"
if ((($ASKPFiles | Measure-Object).Count) -gt 0)
{
	foreach($file in $ASKPFiles)
	{
		$si = $CopyTbl | ? {($_.ASKP.Length -ne 0) -and ($file.Name.Contains($_.ASKP))}
		if ((($si | Measure-Object).Count) -gt 0)
		{
			$CopyPath = $ImpPath + "\" + $si.DIR + "\ASKP"
			Move-Item $file $CopyPath -Force
			if ($Log) {"$(&dmy)`t$($si.DIR)`t- Moved`t->`t$($file.Name) `tto`t$($si.DIR)\ASKP" >> $fl}
		}
	}
	$Cop = $true
}

$XMLUnknownFiles = Get-ChildItem $ImpDir
if ((($XMLUnknownFiles | Measure-Object).Count) -gt 0)
{
	foreach($file in $XMLUnknownFiles)
	{
		$CopyPath = $ImpPath + "\Unknown"
		Move-Item $file $CopyPath -Force
		if ($Log) {"$(&dmy)`tUNKNOWN`t- Moved`t->`t$($file.Name) `tto`tUNKNOWN !!!!!!!" >> $fl}
	}
	$Cop = $true
}

if ($Log -and $Cop) {"---------------------------------------------------------------------------------------------------" >> $fl}