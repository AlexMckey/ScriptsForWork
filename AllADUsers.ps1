[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Add-Type -AssemblyName System.DirectoryServices.Protocols
#Import-Module "d:\Devs\S.DS.P\S.DS.P.psm1"

$searcher = [adsisearcher]"(&(objectclass=user)(objectcategory=person))"
$searcher.SearchRoot = 'LDAP://DC=oduyu,DC=so'
$searcher.PageSize = 1000
$searcher.PropertiesToLoad.AddRange(('samaccountname'))

$query01 = ((Get-QADUser -Filter 'objectCategory -eq "person" -and objectClass -eq "user"' -SearchBase "DC=oduyu,DC=so" -Properties SamAccountName).SamAcountName).Count

Get-Variable query* | Sort name