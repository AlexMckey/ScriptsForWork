dir "h:\Отдел оперативной эксплуатации АСУ\АУДИО-ВИДЕО\" -Recurse -Include "*.mp4" | % {
	$d1 = $_.LastWriteTime.ToString("yyyy_MM")
	$d2 = $_.LastWriteTime.ToString("yyyy_MM_dd")
	$t = $_.LastWriteTime.ToString("HH_00")
	$p = "\\intra-volgd\videoarch$\" + $d1 + "\" + $d2 + "\" + $t
	md $p -Force | out-Null
	mi $_.FullName -Destination $p
}