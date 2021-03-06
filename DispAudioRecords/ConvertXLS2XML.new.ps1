param([string] $MP3Path = "\\intra-volgd\ControlPeregovor")
function ConvertDataToXML([string[]] $data)
{
	$dx = ($data |
	% { $a1 = $_.Split("%");
		$dsp = $a1[0];
		$dd = $a1[1].Split(";")
		$dd |
		? {$_[0] -ne $null} |
		% {
			$a2 = $_.Split("#")
			New-Object PSObject -Property @{
				"Дата" = [DateTime]::Parse($a2[0])
				"Диспетчер" = $dsp
				"Смена" = $a2[1] }
		}
	} |
	Sort Дата |
	Select Дата, Диспетчер, Смена)
	$kv = [int][Math]::Floor(($dx[-1].Дата.Month-1)/3+1)
	$dx | Export-Clixml "$MP3Path\ControlTable.$kv.xml"
}

#Функция чтения Excel-евского файла и преобразования к требуемому нам виду
function GetDataFromXLS([string] $XLSFileName)
{
	#Открываем в excel-е необходимый нам файл
	$objExcel = New-Object -comobject Excel.Application
	$objWorkbook = $objExcel.Workbooks.Open($XLSFileName)
	$objSheet = $objWorkbook.Sheets.Item(1)
	#Начало областей где лежат данные по диспетчерам (r - строка; с - столбец)
	#Ищем области с данными по диспетчерам
	$ar = @()
	for($r=1; $r -le 100; $r++)
	{
		$s1 =  $objSheet.Cells.Item($r,1).Value()
		if ($s1 -eq "1") {$ar += @{c=2;r = $r}}
	}
	#Три области по количеству месяцев в квартале
	#$ar = @{c=2;r=17},@{c=2;r=34},@{c=2;r=51}
	#Цикл по всем этим областям
	$outs = foreach($b in $ar)
	{
		#Цикл по всем диспетчерам в это области
		for($r=0; $r -le 11; $r++)
		{
			#Читаем ФамилиюИО диспетчера и преобразовываем к необходимому виду
			$s1 =  $objSheet.Cells.Item($b.r+$r,$b.c).Value()
			$s = $s1.Replace(".","").Replace(" ","")
			#Разделителем между Диспетчером и другими данными является %
			$s += "%"
			#Цикл по всем проверяющим чтобы собрать все даты когда необходимо проверить записи диспетчера
			for($c=1; $c -le 7; $c++)
			{
				#Если дата имеется
				$v = $objSheet.Cells.Item($b.r+$r,$b.c+$c).Value()
				if ($v -ne $null)
				{
					#Если указана дата для второй смены - то преобразовываем ее к требуемому виду
					$i = $v.IndexOf("-")
					$s2 = if ($i -lt 0) {$v} else {$v.Substring(0,$i) + $v.Substring($i+3)}
					#Удалаем лишние символы
					$s += $s2.Replace(")","").Replace(" (","#").Replace("(","#")
					#Разделителем между датами является ;
					$s += ";"
				}
			}
			#Удалем завершающий разделитель
			$s.TrimEnd(";")
		}
	}
	#Закрываем книгу Excel
	#$objWorkbook.Close()
	$objExcel.Workbooks.Close()
	#Выходим из Excel (вернее даем команду на выход из Excel)
	$objExcel.Quit()
	#обнуляем объект
	$objExcel = $null
	#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
	[gc]::collect()
	[gc]::WaitForPendingFinalizers()
	$outs
}

$xmlf = @()
#Все excel-евские файлы в корневом каталоге
$f = @(dir "$MP3Path\*.xls*")
#Файл должен быть только один
if ($f.Count -gt 0)
{
	$fi = @()
	$fi = if (Test-Path "$MP3Path\ExcelFiles.xml") {Import-Clixml "$MP3Path\ExcelFiles.xml"}
	$fx = @()
	$f | % `
	{
		$fb = $true
		$fl = $_
		$fi | % { if ($_.Name -eq $fl.Name) { $fb = ($_.Length -ne $fl.Length) -or ($_.LastWriteTime -ne $fl.LastWriteTime) }
	}
		if ($fb) { $fx += $fl }
	}
	$fx | % `
	{
		#Читаем данные из файла и преобразовываем в удобный для нас формат
		$data = GetDataFromXLS $_.FullName	
		#Сохраняем данные в XML для дальнейшего использования
		ConvertDataToXML $data
		#Сохраняем информацию о создании файла xml в лог-файл
		[DateTime]::Now.ToString("dd.MM.yyyy HH:mm") + `
		"`t - Создан новый XML файл со сведениями о контроле переговоров" + "`n" `
		>> "$MP3Path\AutoCopyScript.log"
	}
	
	#Сведения о том что файлы были обработаны сохранем в XML-ный файлик в корневой директории
	if ($fx.Count -gt 0)
	{
		$f |
		Select Name, Length, LastWriteTime |
		Export-Clixml "$MP3Path\ExcelFiles.xml"
	}
}