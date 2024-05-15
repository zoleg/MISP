# SETTING UP NECESSARY PARAMETERS
$emailTo = "your to email" # Separate by comma for multiple email addresses
$emailCC = "" # Separate by comma for multiple email addresses
$emailFrom = "your from email"
$emailFromDisplayName = "MISP script executed"
$smtpServer = "smtp.sendgrid.net"
$smtpPort = 587 # 0 = not in use
$enableSSL = $true
$useDefaultCredentials = $false
$username = "apikey" # Provide this when $useDefaultCredentials = $false
$password = "SG.xxxxxxxxx" # Provide this when $useDefaultCredentials = $false
$useHTML = $true
$priority = [System.Net.Mail.MailPriority]::Normal # Low|Normal|High

# ACQUIRING DATA OR INFO, INTERPRET IT AND PROCESS THE RESULT
# Many things you can do here...

# CREATING MESSAGE BODY
# Generally I use HTML for email message body.
$html_var1 = "MISP script executed"
#$html_var2 = $indicator
#$html_var3 = "bla"
$emailBody = $html_var1 #+ $html_var2 + $html_var3

# CREATING EMAIL SUBJECT
# Sometimes I decide the email subject at last, or auto-generate it.
$emailSubject = "MISP for Defender IoCs uploaded successfully"

# SENDING EMAIL...
# You might want to do checking/validation beforehand on necessary variables
$mailMessage = New-Object System.Net.Mail.MailMessage
if ($emailFromDisplayName -eq "" -or $emailFromDisplayName -eq $null) {
    $mailMessage.From = New-Object System.Net.Mail.MailAddress($emailFrom)
} else {
    $mailMessage.From = New-Object System.Net.Mail.MailAddress($emailFrom, $emailFromDisplayName)
}
$mailMessage.To.Add($emailTo)
if ($emailCC -ne "") {
    $mailMessage.CC.Add($emailCC)
}
$mailMessage.Subject = $emailSubject
$mailMessage.Body = $emailBody
$mailMessage.IsBodyHtml = $useHTML
$mailMessage.Priority = $priority

$smtp = $null
if ($smtpPort -ne 0) {
    $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
} else {
    $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer)
}
$smtp.EnableSsl = $enableSSL
$smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network

if ($useDefaultCredentials -eq $true) {
    $smtp.UseDefaultCredentials = $true
} else {
    $smtp.UseDefaultCredentials = $false
    $smtp.Credentials = New-Object System.Net.NetworkCredential($username, $password)
}

try {
    Write-Output ("Sending email notification...")
    $smtp.Send($mailMessage)
    Write-Output ("Sending email notification --> Complete")
} catch {
    Write-Output ("Sending email notification --> Error")
    Write-Warning ('Failed to send email: "{0}"', $_.Exception.Message)
}