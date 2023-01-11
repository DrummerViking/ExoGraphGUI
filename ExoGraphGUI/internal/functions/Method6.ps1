Function Method6 {
    <#
    .SYNOPSIS
    Method to list items in a specific folders in the user mailbox.
    
    .DESCRIPTION
    Method to list items in a specific folders in the user mailbox.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadBasic
    Application: Mail.ReadBasic.All
    
    .PARAMETER folderID
    FolderID value.
    
    .PARAMETER StartDate
    StartDate to search for items.
    
    .PARAMETER EndDate
    EndDate to search for items.

    .PARAMETER MsgSubject
    Optional parameter to search based on a subject text.
    
    .EXAMPLE
    PS C:\> Method6
    
    Method to list items in a specific folders in the user mailbox.

    #>
    [CmdletBinding()]
    Param(
        [String] $folderID,
        [string] $StartDate,
        [string] $EndDate,
        [String] $MsgSubject
    )
    $statusBarLabel.Text = "Running..."
    
    if ( $folderID -ne "" ) {
        Write-PSFMessage -level host -message "current folderID: $folderID, $startdate and $enddate"
        # Creating Filter variables
        $filter = "ReceivedDateTime ge $StartDate and receivedDateTime lt $EndDate"
        if ($MsgSubject -ne "") {
            $filter += " and Subject -eq '$MsgSubject'"
        }
        
        $array = New-Object System.Collections.ArrayList
        $msgs = Get-MgUserMailFolderMessage -UserId $Account -MailFolderId $folderID -Filter $filter | Select-Object subject, @{N = "Sender"; E = { $_.Sender.EmailAddress.Address } }, ReceivedDateTime, isRead
        $null = $array.AddRange($msgs)

        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBarLabel.text = "Ready. Items found: $($array.Count)"
        Write-PSFMessage -Level Output -Message "Task finished succesfully" -FunctionName "Method 6" -Target $email
    }
    else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}