function Send-O365EWSMailMessage {
    <#
.SYNOPSIS
    A function to send mails using the Exchange Web Services API
.DESCRIPTION
    Since Send-MailMessage is considered to be obsolte, 
    this function provides a, pretty much, "drop-in" replacement using the Microsoft Exchange Web Services Managed API.
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage
    https://www.microsoft.com/en-us/download/details.aspx?id=42951

    Obviously the Microsoft Exchange Web Service Managed API has to be installed and you have to use Office365/Exchange.
    The function will fetch the newest installed version if there is more than one on the system (i.e. 1.1 & 2.2)

    Missing Paramaters Send-MailMessage had:

    -DeliveryNotificationOption 
    -Encoding   ->  EWS uses an HTML Body as Default, the encoding has to be set in the HTML Header of the Mail Body. 
                    If no encoding is set UTF8 will be used
    -SmtpServer ->  Well, who would have thought that right?
    -UseSSL     ->  It's 2020 folks, this is HTTPS ONLY
    -Port       ->  Uses TCP443

.PARAMETER Attachment
    Specifies the path and file names of files to be attached to the email message.
.PARAMETER AutoDiscovery 
    An exchange email address that will be used for EWS autodiscovery 
    instead of the hardcoded URI 'https://outlook.office365.com/EWS/Exchange.asmx'
.PARAMETER Subject
    The subject of the mail
.PARAMETER Bcc
    Specifies the email addresses that receive a copy of the mail but are not listed as recipients of the message.
    Enter names (optional) and the email address, such as Name <someone@fabrikam.com>
.PARAMETER Body
    Specifies the content of the email message. 
.PARAMETER BodyAsHtml
    Specifies that the value of the Body parameter contains HTML.
    This defaults to "Text" like Send-MailMessage did.
.PARAMETER Cc
    Specifies the email addresses to which a carbon copy (CC) of the email message is sent.
    Enter names (optional) and the email address, such as Name <someone@fabrikam.com>.
.PARAMETER Credential
    Specifies a user account that has permission to perform this action. The default is the current user.
    Or, enter a PSCredential object, such as one from the Get-Credential cmdlet.
.PARAMETER ExchangeVersion
    String that specifies the [Microsoft.Exchange.WebServices.Data.ExchangeVersion] used in the ExchangeService Constructor 
.PARAMETER From
    The From parameter is NOT required like in Send-MailMessage because EWS will default to the User making the request. 
    This parameter specifies the sender's email address. Enter a name (optional) and email address, such as Name <someone@fabrikam.com>.
.PARAMETER ReplyTo
    Specifies additional email addresses (other than the From address) to use to reply to this message.
    Enter names (optional) and the email address, such as Name <someone@fabrikam.com>.
.PARAMETER SendOnly 
    As Default this function uses SendAndSaveCopy() 
    If SendOnly is set it will just send the mail but not put it in your 'Sent Items' folder. 
.PARAMETER Subject
    The Subject parameter isn't required. This parameter specifies the subject of the email message.
.PARAMETER To
    The To parameter is required. This parameter specifies the recipient's email address. 
    If there are multiple recipients, separate their addresses with a comma (,). 
    Enter names (optional) and the email address, such as Name <someone@fabrikam.com>.
.EXAMPLE
    Send-O365EWSMailMessage -From 'User01@domain.com ' -To 'User02@domain.com' -Subject 'Test mail'
.EXAMPLE
    Send-O365EWSMailMessage -From 'User01@domain.com' -To 'User02@domain.com', 'User03@domain.com' -Subject 'Sending the Attachment' -Body "Forgot to send the attachment. Sending now." -Attachments .\data.csv -Priority High
.EXAMPLE
    Send-O365EWSMailMessage -From 'User01@domain.com' -To 'ITGroup@domain.com' -Cc 'User02@domain.com' -Bcc 'ITMgr@domain.com' -Subject "Don't forget today's meeting!" -Credential domain01\admin01
#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            ValueFromPipeline = $true)]
        [Alias('PsPath')]
        [ValidateScript( { Test-Path -Path $_ })]
        [string[]]
        $Attachments,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AutoDiscovery,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Bcc,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,           

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [Alias('BAH')]
        [switch]
        $BodyAsHtml,     

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Cc,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ExchangeVersion,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $From,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('Normal', 'High', 'Low')]
        [string]
        $Priority = 'Normal', 
        
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ReplyTo,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [bool]
        $SendOnly,  
        
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 1)]
        [Alias('sub')]
        [ValidateNotNullOrEmpty()]
        [String]
        $Subject,  

        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $To
    )
    begin {
        # Gets the most recent version of EWS API that is installed 
        [string]$EWSDLLName = 'Microsoft.Exchange.WebServices.dll'
        [string]$EWSRegPath = Get-ChildItem -ErrorAction SilentlyContinue -Path 'HKLM:\SOFTWARE\Microsoft\Exchange\Web Services' |
            Sort-Object Name -Descending | 
            Select-Object -First 1 -ExpandProperty Name

        [string]$EWSDLLDirectory = $(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$EWSRegPath).'Install Directory'
        [string]$EWSDLLFilePath = $EWSDLLDirectory + $EWSDLLName

        if (Test-Path $EWSDLLFilePath) {
            Import-Module -Name $EWSDLLFilePath
        } else {
            [string]$ExceptionMessage = 'Could not find Microsoft.Exchange.WebServices.dll' + [Environment]::NewLine +
            'Download @ https://www.microsoft.com/en-us/download/details.aspx?id=42951'
            throw [System.IO.FileNotFoundException]::new($ExceptionMessage)
        }
    }
    process {
        # EWS service object                  
        if (!$ExchangeVersion) {
            $ExchVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion
            $ExchService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::New($ExchVersion) 
        } else {
            $ExchService = [Microsoft.Exchange.WebServices.Data.ExchangeService]::New() 
        }

        # Set Credentials
        if ($null -ne $Credential) { 
            # We use NetworkCredential within WebCredentil so we can keep the password as a SecureString
            $ExchCreds = [System.Net.NetworkCredential]::New($Credential.UserName, $Credential.Password)
            $ExchService.Credentials = [Microsoft.Exchange.WebServices.Data.WebCredentials]::New($ExchCreds)
            #$ExchService.Credentials = [Microsoft.Exchange.WebServices.Data.WebCredentials]::New($Credential.UserName, $Credential.GetNetworkCredential().password)
        } else {
            $ExchService.UseDefaultCredentials = $true
        }

        # EWS Endpoint
        if (!$AutoDiscovery) {
            $ExchService.Url = [System.Uri]'https://outlook.office365.com/EWS/Exchange.asmx'  
        } else {
            $ExchService.AutoDiscoverUrl("$AutoDiscovery", { $true })
        }      

        # Create the mail message
        $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::New($ExchService)
        # Add each recipient 
        $To | ForEach-Object { $Message.ToRecipients.Add("$_") | Out-Null }

        # Add all non-mandatory properties
        if ($Attachments) { $Attachments | ForEach-Object { $Message.Attachments.AddFileAttachment($_) } | Out-Null } 
        if ($Body) { $Message.Body = $Body }
        # Add each blind copy recipient    
        if ($Bcc) { $Bcc | ForEach-Object { $Message.BccRecipients.Add("$_") | Out-Null } }
        # If the -BodyAsHtml parameter is not used, send the message as plain text just like Send-MailMessage did
        if (!$BodyAsHtml) { $Message.Body.BodyType = 'Text' } 
        # Add each carbon copy recipient   
        if ($Cc) { $Cc | ForEach-Object { $Message.CCRecipients.Add("$_") | Out-Null } }
        if ($From) { $Message.From = $From }
        if ($Priority) { $Message.Importance = $Priority }
        if ($Subject) { $Message.Subject = $Subject }
        if ($ReplyTo) { $Message.ReplyTo = $ReplyTo }

        try {
            if (!$SendOnly) {
                # Send the message and save a copy in the "Sent Items" folder
                $Message.SendAndSaveCopy()
            } else {
                # Just send the message
                $Message.Send()
            }
        } catch {
            throw $_
        } finally {
            # Uncomment if, for whatever reason, you decide not to use the Network Credential method.
            # https://get-powershellblog.blogspot.com/2017/06/how-safe-are-your-strings.html
            #$ExchService.Credentials = $null
            #$ExchService = $null
            #[System.GC]::Collect()
        }
    }
    end {}
}
Export-ModuleMember -Function 'Send-O365EWSMailMessage'