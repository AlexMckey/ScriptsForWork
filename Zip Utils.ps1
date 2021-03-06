#usage: new-zip c:\demo\myzip.zip2
function New-Zip
{
    param([string]$zipfilename)
	set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
	(dir $zipfilename).IsReadOnly = $false
}
#usage: dir c:\demo\files\*.* -Recurse | add-Zip c:\demo\myzip.zip3
function Add-Zip
{
	param([string]$zipfilename)
	if(-not (test-path($zipfilename)))
	{
		set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(dir $zipfilename).IsReadOnly = $false	
	}
	$shellApplication = new-object -com shell.application
	$zipPackage = $shellApplication.NameSpace($zipfilename)
	foreach($file in $input) 
	{ 
            $zipPackage.CopyHere($file.FullName)
            Start-sleep -milliseconds 500
	}
}
#usage: Get-Zip c:\demo\myzip.zip4
function Get-Zip
{
	param([string]$zipfilename)
	if(test-path($zipfilename))
	{
		$shellApplication = new-object -com shell.application
		$zipPackage = $shellApplication.NameSpace($zipfilename)
		$zipPackage.Items() | Select Path
	}
}
#usage: extract-zip c:\demo\myzip.zip c:\demo\destination
function Extract-Zip
{
	param([string]$zipfilename, [string] $destination)

	if(test-path($zipfilename))
	{	
		$shellApplication = new-object -com shell.application
		$zipPackage = $shellApplication.NameSpace($zipfilename)
		$destinationFolder = $shellApplication.NameSpace($destination)
		$destinationFolder.CopyHere($zipPackage.Items())
	}
}