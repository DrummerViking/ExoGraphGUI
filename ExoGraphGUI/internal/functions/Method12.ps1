Function Method12 {
    <#
    .SYNOPSIS
    Function to inject an email messages through MS Graph.
    
    .DESCRIPTION
    Function to inject an email messages through MS Graph.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadWrite
    Application: Mail.ReadWrite

    .PARAMETER Account
    User's UPN to inject the email message from.

    .PARAMETER ToRecipients
    List of recipients in the "To" list. If ommitted, it will be used the same as the logged on user.

    .PARAMETER NumberOfMessages
    Number of messages to be injected into the Inbox folder. By default is 1.

    .PARAMETER Subject
    Use this parameter to set the subject's text. By default will have: "Test message sent via Graph".

    .PARAMETER Body
    Use this parameter to set the body's text. By default will have: "Test message sent via Graph using Powershell".
    
    .PARAMETER UseAttachment
    Use this switch parameter to add an attachment to sample messages.

    .EXAMPLE
    PS C:\> Method12 -ToRecipients "john@contoso.com"
    Then will send the email message to "john@contoso.com" from the user previously authenticated.

    .EXAMPLE
    PS C:\> Method12 -ToRecipients "julia@contoso.com","carlos@contoso.com" -Subject "Lets meet!"
    Then will send the email message to "julia@contoso.com" and "carlos@contoso.com" and bcc to "mark@contoso.com", from the user previously authenticated.
#>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Cmdletbinding()]
    Param (
        [String] $Account,

        [String[]] $ToRecipients,

        [int] $NumberOfMessages = 1,

        [String] $Subject,

        [String] $Body,

        [Switch] $UseAttachment
    )
    $statusBarLabel.Text = "Running..."

    if ( $ToRecipients -eq "" ) { $ToRecipients = $Account }
    if ( $subject -eq "" ) { $Subject = "Test message injected via Graph" }
    if ( $Body -eq "" ) { $Body = "Test message injected via Graph using ExoGraphGUI tool" }

    # Base mail body Hashtable
    $global:MailBody = @{
        Importance = "Low"
        Sender = @{
            EmailAddress = @{
                Address = $Account
            }
        }
        Subject = $Subject
        Body    = @{
            Content     = $Body
            ContentType = "HTML"
        }
        parentFolderId = (Get-MgUserMailFolder -UserId $Account -MailFolderId "inbox").Id
    }
    
    $attachParams = $null
    if ( $UseAttachment ) {
        # create sample attachment file
        if ( -not(test-path "$env:Temp\SampleFileName.txt") ) {
            $progresBar = New-BTProgressBar -Indeterminate -Status "working"
            New-burntToastNotification -ProgressBar $progresBar -UniqueIdentifier "bar001" -Text "Creating sample file"

            [int]$i = 0
            1..5000 | ForEach-Object {
                $i++
                Write-Progress -activity "Creating sample file. Please wait..." -status "Percent scanned: " -PercentComplete ($i * 100 / 5000) -ErrorAction SilentlyContinue
                "test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool. test file for ExoGraphGUI tool.  " | Add-Content -Path "$env:Temp\SampleFileName.txt" -Force
            }
        }
        $fileContentInBytes = [System.Text.Encoding]::UTF8.GetBytes((Get-Content -path "$env:Temp\SampleFileName.txt"))
        $attachParams = @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name          = "SampleFileName.txt"
            ContentBytes  = [System.Convert]::ToBase64String($fileContentInBytes)
        }
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
    $global:MailBody.Add("ToRecipients", $recipientsList)

    # Making Graph call to inject email message
    try {
        [int]$i = 0
        $progresBar = New-BTProgressBar -Indeterminate -Status "working"
        New-burntToastNotification -ProgressBar $progresBar -UniqueIdentifier "bar001" -Text "Creating sample messages"

        1..$NumberOfMessages | ForEach-Object {
            $i++
            Write-Progress -activity "Creating sample message. $i / $NumberOfMessages" -status "Percent scanned: " -PercentComplete ($i * 100 / $NumberOfMessages) -ErrorAction SilentlyContinue
            $msg = New-MgUserMessage -UserId $Account -BodyParameter $MailBody
            if ( $UseAttachment ) {
                New-MgUserMessageAttachment -UserId $Account -MessageId $msg.Id -BodyParameter $attachParams
            }
            Move-MgUserMessage -UserId $Account -MessageId $msg.id -DestinationId "inbox"
        }
        $statusBarLabel.text = "Ready. Mail injected."
        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 12" -Target $Account
    }
    catch {
        $statusBarLabel.text = "Something failed to inject the email message using graph. Ready."
        Write-PSFMessage -Level Error -Message "Something failed to inject the email message using graph. Error message: $_" -FunctionName "Method 11" -Target $Account
    }
}