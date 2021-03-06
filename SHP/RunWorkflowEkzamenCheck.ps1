param($wfname = "Предупреждение об экзамене", $listname = "График проверки знаний", $url="http://shp-volgd/TechApp/Ekzamen")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists[$listname]
#$list.WorkflowAssociations
$wf = $list.WorkflowAssociations | ? {$_.Name -eq $wfname}
$list.Items | ? {($_["Дата следующей проверки"] -ne $null) -and ((($_["Дата следующей проверки"] - [DateTime]::Today).Days -eq 5) -or (($_["Дата следующей проверки"] - [DateTime]::Today).Days -eq 0))} | % {$site.WorkflowManager.StartWorkflow($_,$wf,$wf.AssociationData,$true)}