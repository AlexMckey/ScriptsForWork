# Параметры: [id_servera,id_состояния] (числа) или [сервер состояние] (строковое представление)
#1. Сервер (можно использовать, как число = идентификатор сервера, так и наименование сервера)
#2. Состояние (можно использовать как число = идентификатор состояния, так и сокращенное наименование состояния)
#
# Сервера:
# 0 = RDUDB
# 1 = RDUARCH
# 2 = RDUARCH2
#
# Состояние:
# 0 = Выключен
# 1 = Включен
# 2 = Потеряна связь

$ParCount = $args.Length
$Debug = $args[-1] -eq "debug"
if ($Debug) {$Line = "-------------------------------------------------------------------------------------------------------------------------------------------"}
if ($Debug) {$DF = "UPS_Event_2_OIK.debug"}
if ($Debug) {"Debug # UPS_Event_2_OIK Report #  Date: $([DateTime]::Now)" > $DF}
if ($Debug) {$Line >>$DF}
if ($Debug) {$ParCount -= 1}
if ($ParCount -eq 0) {exit}
$Params = -1,-1
if ($ParCount -eq 1)
{
  if ($args[0].Length -ne 2)
  {
    if ($Debug) {"Server or State is not assign ==> Exit" >>$DF}
    exit
  }
  $Params = $args[0]
}
else
{
  switch ($args[0])
  {
    RDUDB {$Params[0] = 0}
    RDUARCH {$Params[0] = 1}
    RDUARCH2 {$Params[0] = 2}
  }
  switch ($args[1])
  {
    OFF {$Params[1] = 0}
    ON {$Params[1] = 1}
    LOST {$Params[1] = 2}
  }
}
if ($Debug) {"Params:" >>$DF}
if ($Debug) {$Params >>$DF}
$CK = new -comobject "OIC.DAC"
$CK.Connection.RTDBTaskName = "UPS Event Writer"
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
$EventSender = $CK.OICEvents.EventDispatch
$Event = $EventSender.EventStructure
$Event.EventID = 100002
$Event.CategoryID = 5000
$Event.LevelID = 100
$Event.Flags = 0x40
#$Event.Time = [DateTime]::Now
$Event.Params = $Params
if ($Debug) {"Event: " >> $DF; $Event >> $DF}
$EventSender.Send()
if ($Debug) {"Sending UPS Event" >>$DF}
$CK.Connection.Connected = $false
if ($Debug) {"Sending UPS Event complited" >>$DF}
if ($Debug) {$Line >>$DF}