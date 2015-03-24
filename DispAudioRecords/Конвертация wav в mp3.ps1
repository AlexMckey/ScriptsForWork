param(  [string] $PgvrPath = "w:",
        [string] $ToolPath = "j:\EXE\Script")
Set-Alias Lame "$ToolPath\lame.exe"
dir "$PgvrPath\*" -Recurse -Include "*.wav" | % {
	Lame -f -m m -S $_.FullName $_.FullName.Replace($_.Extension,".mp3")
	ri $_.FullName
}