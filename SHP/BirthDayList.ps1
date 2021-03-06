param(
    #[string[]]$MailFileTo = "sdy@volgograd.so-ups.ru",
    #[string[]]$MailTo = "sdy@volgograd.so-ups.ru",
    [string]$ScriptDir = "D:\Script",
	#[string]$MailTo = "sdy@volgograd.so-ups.ru",
    [string[]]$LogPath = "D:\Script\BirthDayLog.txt"
)
# Подгружаем сборку Sharepoint
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Начали" >> $LogPath
"-"*90 >> $LogPath
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Подгрузил сборку Sharepoint;">> $LogPath
"-"*90 >> $LogPath
# Получаем список именинников
$listname = "Дни рождения"
$url="http://shp-volgd/ShareDoc"
$site = new-object Microsoft.SharePoint.SPSite("$url")
$web = $site.OpenWeb()
$list = $web.Lists[$listname]
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Получил список именинников;">> $LogPath
"-"*90 >> $LogPath
# Поучаем список адресатов
$listnamedest = "Персонал"
$urldest="http://shp-volgd/OtdelDocs/APO"
$sitedest = new-object Microsoft.SharePoint.SPSite("$urldest")
$webdest = $sitedest.OpenWeb()
$listdest = $webdest.Lists[$listnamedest]
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
" Получил список адресатов;">> $LogPath
"<>"*45 >> $LogPath
# Здесь формируем список рассылки
$psbdlist = $list.Items | % {New-Object PSObject -Property @{"FIO"=$_["ФИО"]; "Date"=$_["Дата"].ToShortDateString(); "JOB"=$_["Должность"]}}
$arrdest = $listdest.Items | ? {$_["Список рассылки"] -eq $true} | % {$_["Адрес электронной почты"]}
$MailTo = $arrdest -join ", "
[DateTime]::Today.time >> $LogPath
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Сформировал список рассылки:">> $LogPath
$arrdest >> $LogPath
"<>"*45 >> $LogPath
# уже дополнение
$bddate=$null
$bodylist = $null
# Если сегодня начало месяца делаем рассылку на текущий месяц
$maxname=0
if ([DateTime]::Today.Day -eq 1)
{
$thismonthbd = $psbdlist | ? {([DateTime]::Parse($_.Date).Month) -eq [DateTime]::Today.month} | % {New-Object PSObject -Property @{"IFIO"=$_.FIO;"IDATE"=$_.DATE;"IJOB"=$_.JOB;"YEARS"=([DateTime]::Today.year-[DateTime]::Parse($_.Date).year)}}
$thismonthbd | % {if ($_.IFIO.Length -ge $maxname){$maxname=$_.IFIO.Length}}
#$bddate= $psbdlist | ? {([DateTime]::Parse($_.Date).Month-1) -eq $nowmonth} | % {@([DateTime]::Parse($_.Date).year)}
$bodylist = $thismonthbd | % {"`n" + $_.IJOB +"`n`n`t"+$_.IFIO + "," + " " * ($maxname - $_.IFIO.Length) +"`t "+ [DateTime]::Parse($_.IDate).tostring("d MMMM") + "`t исполнится:`t"+$_.years+"`n`n"+"-"*100+"`n"}
$body = "`t`tВ этом месяце свой день рождения празднуют: `n"+"-"*100+"`n $bodylist"
if ($thismonthbd -ne $null)
{
& "$ScriptDir\Send-SMTPMail.ps1" -To $MailTo -Subject "Уведомление о Днях Рождения" -Body $Body -Server 'ex-oduyu-cas.oduyu.so' -From 'volga_rdu@volgograd.so-ups.ru'
}
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Отправил сообщения по электронной почте о ДР в этом месяце у:" >> $LogPath
$thismonthbd | % {($_.IFIO, $_.IDATE, $_.IJOB, $_.YEARS)>> $LogPath}
"<>"*45 >> $LogPath
}
# Если ДР сегодня делаем рассылку
$thisdaybd=$null
$thisdaybd = $psbdlist | ? {(([DateTime]::Parse($_.Date).Month) -eq [DateTime]::Today.month) -and (([DateTime]::Parse($_.Date).day) -eq [DateTime]::Today.day)} | % {@{"IFIO"=$_.FIO; "IJOB"=$_.JOB; "YEARS"=([DateTime]::Today.year - [DateTime]::Parse($_.Date).year)}}
if ($thisdaybd -ne $null)
{
$thisdaybd | % {if ($_.IFIO.Length -ge $maxname){$maxname=$_.IFIO.Length}}
$bodylist = $thisdaybd | % {"`n" + $_.IJOB +"`n`n " + $_.IFIO+ " " * ($maxname - $_.IFIO.Length)+",`t"+"сегодня ему(ей) исполняется:"+"`t"+$_.years+"`n`n"}
$body = "`t`tСегодня свой день рождения празднуют: `n"+"-"*100+"`n $bodylist"
& "$ScriptDir\Send-SMTPMail.ps1" -To $MailTo -Subject "Уведомление о Днях Рождения" -Body $Body -Server 'ex-oduyu-cas.oduyu.so' -From 'volga_rdu@volgograd.so-ups.ru'
get-date -Format 'dd.MM.yyyy HH:mm:ss' >> $LogPath
"Отправил сообщения по электронной почте о ДР сегодня у:">> $LogPath
$thisdaybd | % {($_.IFIO,  $_.IJOB, $_.YEARS)>> $LogPath}
"<>"*45 >> $LogPath
}
"-"*90 >> $LogPath
" "*90 >> $LogPath
" "*90 >> $LogPath
