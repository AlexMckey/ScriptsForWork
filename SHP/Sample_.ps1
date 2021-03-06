#param($url="http://shp-volgd/ISS")
$url="http://shp-volgd/ISS"
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists["Документы"]
$wf = $list.WorkflowAssociations | ? {$_.Name -eq "Синхронизация имени и названия документа"}
#$i1 = $list.Items | ? {$_.DisplayName -like "Регламент обращения к общему ресурсу (диску Н)_2008"}
#$i1["Название"] = $i1.DisplayName
#$list.Items | ? {$_["Дней до срока"] -ge 0} | % {$site.WorkflowManager.StartWorkflow($_,$wa1,$wa1.AssociationData,$true)}
$list.Items | ? {$_["Название"] -ne $_.DisplayName} | % {$_["Название"] = $_.DisplayName; $_.Update()}
#$list.Items | ? {$_["Название"] -ne $_.DisplayName -and $_["Подразделение"] -eq "СЭР"} | % {$_.DisplayName.Replace("  "," ")}
#$list.Items | ? {$_["Название"] -ne $_.DisplayName -and $_["Подразделение"] -eq "СЭР"} | % {$_.DisplayName.Replace("  "," "); $_["Название"] = $_.DisplayName}
$list.Items | ? {$_["Для подразделений"] -like "*Руководство*"; $_["Для подразделений"].ToString()}