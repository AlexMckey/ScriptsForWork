$what="ОДСИТ"
$to="ООЭ АСУ"
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$u0="http://shp-volgd/ISS"
$s0 = new-object Microsoft.SharePoint.SPSite($u0)
$w0 = $s0.OpenWeb()
$l0 = $w0.Lists["Документы"]
$l0.Items | ? {$_["Для подразделений"] -like "*$what*"} | % {$_["Для подразделений"] = $_["Для подразделений"].Replace(";#$what",";#$to"); $_.Update()}
$i0 = $l0.Items | ? {$_["Подразделение"] -eq $what}
$i0 | % {$_["Подразделение"] = $to; $_.Update()}
$u1 = "http://shp-volgd/ShareDoc"
$s1 = new-object Microsoft.SharePoint.SPSite($u1)
$w1 = $s1.OpenWeb()
$l1 = $w1.Lists["Организационно-распорядительные документы подразделений"]
$i1 = $l1.Items | ? {$_["Подразделение"] -eq $what}
$i1 | % {$_["Подразделение"] = $to; $_.Update()}
$u3="http://shp-volgd"
$s3 = new-object Microsoft.SharePoint.SPSite($u3)
$w3 = $s3.OpenWeb()
$l3 = $w3.Lists["Персонал"]
$i3 = $l3.Items | ? {$_["Подразделение"] -eq $what}
$i3 | % {$_["Подразделение"] = $to; $_.Update()}