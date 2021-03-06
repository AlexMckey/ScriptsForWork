param($url="http://shp-volgd/ISS")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists["Документы"]
if ($site -eq $null -or $web -eq $null -or $list -eq $null)
{
  "ScriptRes:Bad:NoConn"
}
$cnt = 0
$list.Items | ? {$_["Название_"] -ne $_.DisplayName} | % {$_["Название_"] = $_.DisplayName; $_.Update(); $cnt+=1}
if ($cnt -eq 0)
{
  "ScriptRes:Ok:NotNeed"
}
else
{
  "ScriptRes:Ok:$Cnt"
}