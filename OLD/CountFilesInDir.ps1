function MakeDir([String] $dir)
{
  if (Test-Path $dir -PathType Container)
    {$dir+"\*"}
  else
    {(split-path $dir) + "\*"}
}
$dir = MakeDir($args[0])
$inc = $args[1..$args.Length]
$files = dir $dir -include $inc
$cnt = &{if ($files -eq $null) {0} else {$files.Length}}
if ($cnt -eq 0)
  {"ScriptRes:Ok:0"}
else
  {"ScriptRes:Bad:$cnt"}