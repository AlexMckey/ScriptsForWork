Function UploadToSPDocLib($LocalPath,$spDocLibPath) 

{ 

$UploadFullPath = $spDocLibPath + $(split-path -leaf $LocalPath) 

$WebClient = new-object System.Net.WebClient 

$WebClient.Credentials = [System.Net.CredentialCache]::DefaultCredentials 

$WebClient.UploadFile($UploadFullPath, "PUT", $LocalPath) 

}