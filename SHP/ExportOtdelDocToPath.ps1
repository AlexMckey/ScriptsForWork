param($otdel,$path,$url="http://shp-volgd/ISS")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists["Документы"]
$OnePath = $path -eq $null
# $WebClient = new-object System.Net.WebClient
# $WebClient.Credentials = [System.Net.CredentialCache]::DefaultCredentials
$list.Items | 
? {$_["Для подразделений"] -like "*$otdel*" -and $_["Путь экспорта"] -ne $null} |
% {$p = if ($OnePath) {$_["Путь экспорта"]} else {$path}; [System.IO.File]::WriteAllBytes($p + "\" + $_.Name, $_.File.OpenBinary())}
# % {$WebClient.DownloadFile($url+"/"+$list.RootFolder+"/"+$i.Name,$path+"\"+$i.Name)}