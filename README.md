# Send-O365EWSMailMessage
## Replacement for the Send-MailMessage Cmdlet in Powershell

Since Send-MailMessage is considered to be obsolte, 
this function provides a, pretty much, "drop-in" replacement using the Microsoft Exchange Web Services Managed API.

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage
https://www.microsoft.com/en-us/download/details.aspx?id=42951

Obviously the Microsoft Exchange Web Service Managed API has to be installed and you have to use Office365/Exchange.
The function will fetch the newest installed version if there is more than one on the system (i.e. 1.1 & 2.2).

## Missing Paramaters Send-MailMessage had

### DeliveryNotificationOption 
Might be added in future commits.

### Encoding 
EWS uses an HTML Body as Default, the encoding has to be set in the HTML Header of the Mail Body. 
If no encoding is set (or the Body is of BodyType `Text`, UTF8 will be used.

### SmtpServer 
Well, who would have thought that right?

### UseSSL 
It's 2020 folks, this is HTTPS ONLY.

### Port 
Like I said this is HTTPS/TCP443 only.