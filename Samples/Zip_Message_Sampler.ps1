$sender = sender@host.com
$recipient = recipient@host.com
$server = mail.host.com
$targetFolder = c:\MyFolder
$file = c:\MyZipFile.zip

if ( [System.IO.File]::Exists($file) )
{
  remove-item -force $file
}

gi $targetFolder | out-zip $file $_
$subject = "Sending a File " + [System.DateTime]::Now
$body = "I'm sending a file!"
$msg = new-object System.Net.Mail.MailMessage $sender, $recipient, $subject, $body
$attachment = new-object System.Net.Mail.Attachment $file
$msg.Attachments.Add($attachment)
$client = new-object System.Net.Mail.SmtpClient $server
$client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$client.Send($msg)