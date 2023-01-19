Function Method11 {
    <#
    .SYNOPSIS
    Function to send email messages through MS Graph.
    
    .DESCRIPTION
    Function to send email messages through MS Graph.
    Module required: Microsoft.Graph.Users.Actions
    Scope needed:
    Delegated: Mail.Send
    Application: Mail.Send

    .PARAMETER Account
    User's UPN to send the email message from.

    .PARAMETER ToRecipients
    List of recipients in the "To" list. This is a Mandatory parameter.
    
    .PARAMETER CCRecipients
    List of recipients in the "CC" list. This is an optional parameter.

    .PARAMETER BccRecipients
    List of recipients in the "Bcc" list. This is an optional parameter.

    .PARAMETER Subject
    Use this parameter to set the subject's text. By default will have: "Test message sent via Graph".

    .PARAMETER Body
    Use this parameter to set the body's text. By default will have: "Test message sent via Graph using Powershell".
    
    .EXAMPLE
    PS C:\> Method11 -ToRecipients "john@contoso.com"
    Then will send the email message to "john@contoso.com" from the user previously authenticated.

    .EXAMPLE
    PS C:\> Method11 -ToRecipients "julia@contoso.com","carlos@contoso.com" -BccRecipients "mark@contoso.com" -Subject "Lets meet!"
    Then will send the email message to "julia@contoso.com" and "carlos@contoso.com" and bcc to "mark@contoso.com", from the user previously authenticated.
#>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Cmdletbinding()]
    Param (
        [String] $Account,

        [parameter(Mandatory = $true)]
        [String[]] $ToRecipients,

        [String[]] $CCRecipients,

        [String[]] $BccRecipients,

        [String] $Subject,

        [String] $Body
    )
    $statusBarLabel.Text = "Running..."

    if ( $subject -eq "" ) { $Subject = "Test message sent via Graph" }
    if ( $Body -eq "" ) { $Body = "Test message sent via Graph using ExoGraphGUI tool" }

    # Base mail body Hashtable
    $global:MailBody = @{
        Message         = @{
            Subject = $Subject
            Body    = @{
                Content     = $Body
                ContentType = "HTML"
            }
        }
        savetoSentItems = "true"
    }

    # looping through each recipient in the list, and adding it in the hash table
    $recipientsList = New-Object System.Collections.ArrayList
    foreach ( $recipient in ($ToRecipients.split(",").Trim()) ) {
        $recipientsList.add(
            @{
                EmailAddress = @{
                    Address = $recipient
                }
            }
        )
    }
    $global:MailBody.Message.Add("ToRecipients", $recipientsList)

    # looping through each recipient in the CC list, and adding it in the hash table
    if ( $CCRecipients -ne "" ) {
        $ccRecipientsList = New-Object System.Collections.ArrayList
        foreach ( $cc in $CCRecipients.split(",").Trim()) {
            $null = $ccRecipientsList.add(
                @{
                    EmailAddress = @{
                        Address = $cc
                    }
                }
            )
        }
        $MailBody.Message.Add("CcRecipients", $ccRecipientsList)
    }

    # looping through each recipient in the Bcc list, and adding it in the hash table
    if ( $BccRecipients -ne "" ) {
        $BccRecipientsList = New-Object System.Collections.ArrayList
        foreach ( $bcc in $BccRecipients.split(",").Trim()) {
            $null = $BccRecipientsList.add(
                @{
                    EmailAddress = @{
                        Address = $bcc
                    }
                }
            )
        }
        $MailBody.Message.Add("BccRecipients", $BccRecipientsList)
    }

    # Making Graph call to send email message
    try {
        Send-MgUserMail -UserId $Account -BodyParameter $MailBody -ErrorAction Stop
        $statusBarLabel.text = "Ready. Mail sent."
        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 11" -Target $Account
    }
    catch {
        $statusBarLabel.text = "Something failed to send the email message using graph. Ready."
        Write-PSFMessage -Level Error -Message "Something failed to send the email message using graph. Error message: $_" -FunctionName "Method 11" -Target $Account
    }
}