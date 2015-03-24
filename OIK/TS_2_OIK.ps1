# Параметры:
#1. Номера ТС
#2. Значение
$Debug = $args[-1] -eq "debug"
if ($Debug) {$Line = "-------------------------------------------------------------------------------------------------------------------------------------------"}
if ($Debug) {$DF = "TS_2_OIK.debug"}
if ($Debug) {"Debug # TS_2_OIK Report #  Date: $([DateTime]::Now)" > $DF}
if ($Debug) {$Line >>$DF}
if (($Debug -and ($args.Length -eq 1)) -or ($args.Length -eq 0) -or ($args -eq $null)) {exit}
$TS = $args[0]
if (($args[1] -eq "debug") -or ($args.Length -eq 1)) {$Val = 0} else {$Val = $args[1]}
if ($Debug) {"Input Params:" >> $DF}
if ($Debug) {$TS >>$DF}
if ($Debug) {$Val >>$DF}
if ($TS -eq $null)
{
  if ($Debug) {"TS List is null ==> Exit" >>$DF}
  exit
}
$CK = new -comobject "OIC.DAC"
$CK.Connection.RTDBTaskName = "TS Writer"
$CK.Connection.Alias = "Волгоградское РДУ\"
$CK.Connection.Connected = $true
if ($CK.Connection.Connected -eq $false)
{
  if ($Debug) {"OIK not connected ==> Exit" >>$DF}
  exit
}
else
{
  if ($Debug) {"OIK connected" >>$DF}
}
if ($Debug) {$Line >>$DF}
$Req = $CK.OIRequests.Add()
$TS | % `
{
  $Item = $Req.AddOIRequestItem()
  $Item.KindRefresh = 4 #WriteToRTDB
  $Item.DataSource = "S"+$_
  $Item.DataValue = $Val
  $Item.Sign = 0x40
  if ($Debug) {"ReqItem: " >> $DF; $Item >> $DF}
}
$Req.Start()
if ($Debug) {$Line >>$DF}
if ($Debug) {"Start write TS" >>$DF}
#if ($Req.WaitComplete(([DateTime]::Now).AddSeconds(1)))
$Req.Stop()
if ($Debug) {"Stop write TS" >>$DF}
$CK.Connection.Connected = $false
if ($Debug) {"Writing TS complited" >>$DF}
if ($Debug) {$Line >>$DF}