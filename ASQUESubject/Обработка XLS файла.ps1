$Flpath = "\\as2-volgd\ingis\EXE\Script\SubjCopyTbl"
$f = Get-ChildItem "$Flpath.xls*"
if ($f -ne $null)
{
	$fi = if (Test-Path "ExcelFile.xml") {Import-Clixml "ExcelFile.xml"}
	if (($fi -eq $null) -or ($f.LastWriteTime -ne $fi.LastWriteTime))
	{
		# функция обработки XLS файла и сохранения данных CSV
		
		#$xlCSV = 6
		#$Workbook.SaveAs("$Flpath.csv",$xlCSV) 
		
		$f |
		Select Name, LastWriteTime |
		Export-Clixml "ExcelFile.xml"
	}
}