function OtherUploadFile($LocalPath, $SitePath)
{
    $SiteUrl = $SitePath.Substring(0,$SitePath.IndexOf("/",$SitePath.IndexOf("//")+2))
    $FullSitePath = $SitePath
    #if (not $SitePath.EndsWith("/")) {$FullSitePath += "/"}
    if ($SitePath[-1] -ne "/") {$FullSitePath += "/"}
    $FileName = split-path -leaf $LocalPath
    $FullSitePath += $FileName
    $FileName = (split-path -leaf $FileName).Substring(0,$FileName.LastIndexOf("."))
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
    $site = new-object Microsoft.SharePoint.SPSite($SitePath)
    $web = $site.OpenWeb()
    $item = $web.Files.Add($FullSitePath,[System.IO.File]::ReadAllBytes($LocalPath)).Item
    $item["Название"] = $FileName
    $item.Update()
}