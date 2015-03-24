$cnt = 0
if ($args[1] -eq $null)
{
  $cnt = (dir $args[0]).Length
}
else
{
  foreach ($file in dir $args[0] -name)
  {
    foreach ($mask in $args[1])
    {
      if ($file -like $mask) {$cnt += 1}
    }
  }
}
if ($cnt -eq 0)
  {"ScriptRes:Ok:0"}
else
  {"ScriptRes:Bad:$cnt"}