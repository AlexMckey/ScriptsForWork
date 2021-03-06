param([string] $ImpPath = "k:\Import", [string] $ProgPath = "j:\EXE", [string] $ArchDir = "j:\EXE\Script")

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
		$file.Delete()
	}
	Pop-Location
}

$ImpDir = MakeDir($ImpPath)
# разархивируем архивы
Get-ChildItem $ImpPath "*.zip" | UnzipFiles $ImpPath
Get-ChildItem $ImpPath "*.rar" | UnrarFiles $ImpPath
Get-ChildItem $ImpPath "*.7z" | Un7zFiles $ImpPath
# удаляем мусор
Remove-Item $ImpDir -Include "*.html,*.log,*80030*.xml,*80040*.xml,*.jpg,*80040*.xml,*.xls,*.xlsx".Split(",") -Exclude "Script.log"
# проверяем какие программы запушены
$Programs = Get-Process | ? {$_.ProcessName.Contains("INGIS")} | % {$_.ProcessName.Split("_")[0]}

# проверяем файлы формата XML80020
$XML80020Files = Get-ChildItem $ImpDir -Include "*80020*.xml"
$XML80020Files | ? { ($_.Name -like "80020_4716016979*") -and ($_.Length -le 2Mb) } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_3435098928*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_2460066195*") -and ($_.Length -ge 20kb) } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_7709331020*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_3437000021*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_3435000467*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_3435900186*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_6612005052*") } | Copy-Item -Destination "\\appsrv-volgd\c$\Program Files\LocalEscortService\InputFiles\"
$XML80020Files | ? { ($_.Name -like "80020_7709331020*") } | Move-Item -Destination "\\as2-volgd\ingis\ImpExp\TNS" -Force
#if (((($XML80020Files | Measure-Object).Count) -gt 0) -and ($Programs -notcontains "XML80020"))
#    { & "$ProgPath\XML80020_2_INGIS.exe" "auto" }

# проверяем файлы формата XML51070
#$XML51070Files = Get-ChildItem $ImpDir -Include "*51070*.xml"
#if (((($XML51070Files | Measure-Object).Count) -gt 0) -and ($Programs -notcontains "XML51070"))
#    { & "$ProgPath\XML51070_2_INGIS.exe" "auto" }

# проверяем файлы формата ASKP
#$ASKPFiles = Get-ChildItem  $ImpDir -Include "vtz*","az*"
#if (((($ASKPFiles | Measure-Object).Count) -gt 0) -and ($Programs -notcontains "ASKP"))
#    { & "$ProgPath\ASKP_2_INGIS.exe" "auto" }

# проверяем файлы формата XML
#$OtherXMLFiles = Get-ChildItem $ImpDir -Include "*.xml" -Exclude "vtz*","az*","*51070*.xml","*80020*.xml","*.html","*.log","*80030*.xml","*80040*.xml","*.xls","*.xlsx"
#if ((((Get-ChildItem $AODir -Include "*.xml" | Measure-Object).Count) -gt 0) -and ($Programs -notcontains "XML51070"))
#if (((($OtherXMLFiles | Measure-Object).Count) -gt 0) -and ($Programs -notcontains "XML51070"))
#    { & "$ProgPath\XML51070_2_INGIS.exe" "auto" }