########################################################
#
# out-zip.ps1
#
# Usage:
# 
# To zip up some files:
# ls c:\source\*.txt | out-zip c:\target\archive.zip $_
#
# To zip up a folder:
# gi c:\source | out-zip c:\target\archive.zip $_
########################################################

$path = $args[0]
$files = $input
  
if (-not $path.EndsWith('.zip')) {$path += '.zip'} 

if (-not (test-path $path)) { 
  set-content $path ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) 
} 

$ZipFile = (new-object -com shell.application).NameSpace($path) 
$files | foreach {$zipfile.CopyHere($_.fullname)}