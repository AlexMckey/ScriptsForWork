$url="http://shp-volgd/ISS"
$what="ОДСИТ"
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists["Документы"]
$list.Items | ? {$_["Для подразделений"] -like "*$what*"} | % {$_["Для подразделений"] = $_["Для подразделений"].Replace(";#$what",";#ООЭ АСУ"); $_.Update()}
$lc = $list.Items | ? {$_["Подразделение"] -like "*$what*"}
$lc | % {$_["Подразделение"] = $_["Подразделение"].Replace("$what","ООЭ АСУ"); $_.Update()}
$u1 = "http://shp-volgd/ShareDoc"
$s1 = new-object Microsoft.SharePoint.SPSite("$u1")
$w1 = $s1.OpenWeb()
$l1 = $w1.Lists["Организационно-распорядительные документы подразделений"]
$i1 = $l1.Items | ? {$_["Подразделение"] -like "*$what*"}
$i1 | % {$_["Подразделение"] = $_["Подразделение"].Replace("$what","ООЭ АСУ"); $_.Update()}