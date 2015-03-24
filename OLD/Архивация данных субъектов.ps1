param(  [string] $ImpPath = "k:",
        [string] $ArchPath = "j:\Archive\Subject",
		[string] $ToolPath = "j:\EXE\Script")

function GetLabel([IO.FileInfo] $fi)
{
	return ($fi.Name.Split("_")[2]).Substring(0,6)
}

function ArchFiles([string] $lbl)
{
	$archfile = Join-Path $archdir $lbl.Substring(0,4)
	if (-not (Test-Path $archfile)) { md $archfile > null}
	Set-Alias RAR "$ToolPath\Rar.exe"
	RAR m -m5 -ep -inul -y "$archfile\$lbl" "$filesdir\*$lbl*.xml"
	"    Year: " + $lbl.Substring(0,4) + " Month: " + $lbl.Substring(4) + " ArchFile: " + $archfile + "\$lbl.rar  - Complite"
}

function DoSubject([IO.DirectoryInfo] $SubjDir)
{
	"Subject: " + $SubjDir.Name
	$files = $SubjDir.GetFiles("*80020*.xml") | ? { $_.Name -notlike "*$tds*"}
	if (($files -eq $null) -or ($files.Count -eq 0)) {return}
	$lbls = $files | % {GetLabel($_.FullName)} | select -Unique
	$archdir = Join-Path $ArchPath $SubjDir.Name
	$filesdir = $SubjDir.FullName
	$lbls | % { ArchFiles($_) }
}

#$ImpDir = MakeDir($ImpPath)

$td = [DateTime]::Today
$tds = $td.ToString("yyyyMM")

$dirs = Get-ChildItem $ImpPath | ? { $_ -notlike "*Import*"}
$dirs | % { DoSubject($_) }