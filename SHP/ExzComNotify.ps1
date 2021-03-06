$url="http://shp-volgd/TechApp/Ekzamen"
$ComListName = "Экзаменационные комиссии"
$UserListName = "Весь персонал ВРДУ"
$ExzListName = "График проверки знаний"
$ScriptPath="D:\Script"

function SendNotifyUser($un, $ud, $comFIO)
{
    $uc = $UserList | ? {$_.ID -eq $comFIO.split(";")[0]}
    $toM = $uc.Email
    $ds = $d.ToString('dd.MM.yyyy')
    $body = "`t У В Е Д О М Л Е Н И Е !!! `n`n Подходит срок проверки знаний: `n`n Работник - $un `n Дата предстоящей проверки - $ds"
    & "$ScriptPath\Send-SMTPMail.ps1" -To $toM -Body $Body -Subject "Уведомление об экзаменах" -Server "ex-oduyu-cas.oduyu.so" -From 'iss@volgograd.so-ups.ru'
}

function SendNotifyCom($uFIO, $d, $com)
{
    $u = $UserList | ? {$_.ID -eq $uFIO.split(";")[0]}
    $ucs = $ComList | ? {($_["Проверяющая комиссия"] -eq $com) -and ($_["Уведомление"])}
    $ucs | % {SendNotifyUser $u.Name $d $_["ФИО"]}
}

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$ComList = $web.Lists[$ComListName].Items
$UserList = $web.SiteGroups[$UserListName].Users
$ExzList = $web.Lists[$ExzListName].Items
$us = $ExzList | ? {($_["Дата следующей проверки"] -ne $null) -and (($_["Дата следующей проверки"] - [DateTime]::Today).Days -eq 3)}
$us | % {SendNotifyCom $_["ФИО"] $_["Дата следующей проверки"] $_["Проверяющая комиссия"]}