Function Get-UserInboxRule {
    <#
    .SYNOPSIS
    Method to get user's Inbox Rules.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: MailboxSettings.Read
    Application: MailboxSettings.Read
    
    .DESCRIPTION
    Method to get user's Inbox Rules.
    
    .PARAMETER Account
    User's UPN to get mail folders from.
    
    .EXAMPLE
    PS C:\> Get-UserInboxRule
    Method to get user's Inbox Rules.

    #>
    [CmdletBinding()]
    param(
        [String] $Account
    )
    $statusBarLabel.Text = "Running..."

    $array = New-Object System.Collections.ArrayList
    $rules = Get-MgUserMailFolderMessageRule -UserId $Account -MailFolderId "Inbox"
    foreach ( $rule in $rules ) {
        $output = $rule | Select-Object DisplayName, HasError, IsEnabled, IsReadOnly, Sequence
        $array.Add($output)
        Write-PSFMessage -Level Verbose -Message $output -FunctionName "Method6" -Target $Account
    }
    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready..."
    Write-PSFMessage -Level Host -Message "Succesfully retrieved inbox rules." -FunctionName "Method 6" -Target $Account
}