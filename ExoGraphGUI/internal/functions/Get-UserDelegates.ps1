﻿Function Get-UserDelegates {
    <#
    .SYNOPSIS
    Get user's Delegates information
    
    .DESCRIPTION
    Get user's Delegates information
    Module required: Microsoft.Graph.Authentication
    Scope needed:
    Delegated: MailboxSettings.Read
    Application: MailboxSettings.Read

    .PARAMETER Account
    User's UPN to get delegate settings from.
    
    .EXAMPLE
    PS C:\> Get-UserDelegates
    Get user's Delegates information

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "")]
    [CmdletBinding()]
    param(
        $Account
    )
    $statusBarLabel.Text = "Running..."
    $txtBoxResults.Text = "This function is still under construction."
    
    #TODO
    $response = Invoke-MgGraphRequest -Method get -Uri https://graph.microsoft.com/v1.0/users/$Account/mailboxSettings
    $response["delegateMeetingMessageDeliveryOptions"]

    $dgResults.Visible = $False
    $txtBoxResults.Visible = $True
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready."
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 10" -Target $Account
}