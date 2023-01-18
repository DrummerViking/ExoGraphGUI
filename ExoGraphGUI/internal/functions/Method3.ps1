Function Method3 {
    <#
    .SYNOPSIS
    Method to list items in a specific folders in the user mailbox.
    
    .DESCRIPTION
    Method to list items in a specific folders in the user mailbox.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadBasic
    Application: Mail.ReadBasic.All

    .PARAMETER Account
    User's UPN to get mail messages from.
    
    .PARAMETER folderID
    FolderID value to get mail messages from.
    
    .PARAMETER StartDate
    StartDate to search for items.
    
    .PARAMETER EndDate
    EndDate to search for items.

    .PARAMETER MsgSubject
    Optional parameter to search based on a subject text.
    
    .EXAMPLE
    PS C:\> Method3
    
    Method to list items in a specific folders in the user mailbox.

    #>
    [CmdletBinding()]
    Param(
        [String] $Account,
        [String] $folderID,
        [string] $StartDate,
        [string] $EndDate,
        [String] $MsgSubject
    )
    $statusBarLabel.Text = "Running..."
    
    if ( $folderID -ne "" ) {
        # Creating Filter variables
        $filter = $null
        if ($MsgSubject -ne "") {
            $filter = "Subject eq '$MsgSubject'"
        }
        
        $array = New-Object System.Collections.ArrayList
        $msgs = Get-MgUserMailFolderMessage -UserId $Account -MailFolderId $folderID -Filter $filter -All | Where-Object {$_.ReceivedDateTime -ge $StartDate -and $_.ReceivedDateTime -lt $EndDate} | Select-Object subject, @{N = "Sender"; E = { $_.Sender.EmailAddress.Address } }, ReceivedDateTime, isRead
        $null = $msgs | ForEach-Object { $array.Add($_) }
        
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBarLabel.text = "Ready. Items found: $($array.Count)"
        Write-PSFMessage -Level Output -Message "Task finished succesfully" -FunctionName "Method 3" -Target $Account
    }
    else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
        $statusBarLabel.text = "Method 6 finished with warnings/errors"
    }
}