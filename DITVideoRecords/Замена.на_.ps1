# замена в имени директории "." на "_" 
param([string] $root="b:\VideoObhodODSIT\")
Get-ChildItem $root | ? {$_.PSIsContainer -and $_.Name.contains(".")} | % {Rename-Item -Path $_.FullName -NewName ($_.Parent.FullName + "\" + $_.Name.replace(".","_"))}