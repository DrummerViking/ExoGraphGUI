Function Method10 {
    <#
    .SYNOPSIS
    Method to get user's OOF Settings.
    Module required: Microsoft.Graph.Authentication
    Scope needed:
    Delegated: MailboxSettings.Read
    Application: MailboxSettings.Read

    .DESCRIPTION
    Method to get user's OOF Settings.
    
    .PARAMETER Account
    User's UPN to get OOF settings from.

    .EXAMPLE
    PS C:\> Method10
    Method to get user's OOF Settings.

    #>
    [CmdletBinding()]
    param(
        [String] $Account
    )
    $statusBarLabel.Text = "Running..."

    $response = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$Account/mailboxSettings/automaticRepliesSetting"

    $array = New-Object System.Collections.ArrayList
    $output = $response | Select-Object `
        @{ Name = "Status" ; Expression = { $response["Status"] } }, `
        @{ Name = "ExternalAudience" ; Expression = { $response["externalAudience"] } }, `
        @{ Name = "StartTime" ; Expression = { $response["scheduledStartDateTime"].DateTime.ToString("yyyy/MM/dd HH:mm:ss") } }, `
        @{ Name = "EndTime"   ; Expression = { $response["scheduledEndDateTime"].DateTime.ToString("yyyy/MM/dd HH:mm:ss") } }, `
        @{ Name = "InternalReplyMessage" ; Expression = { $response["InternalReplyMessage"] } }, `
        @{ Name = "ExternalReplyMessage" ; Expression = { $response["ExternalReplyMessage"] } }
    $array.Add($output)

    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready..."
    Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 10" -Target $Account
}