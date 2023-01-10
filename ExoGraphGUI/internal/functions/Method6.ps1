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
    
    .EXAMPLE
    PS C:\> Method6
    Method to list items in a specific folders in the user mailbox.

    #>
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBarLabel.Text = "Running..."
    
    if ( $txtBoxFolderID.Text -ne "" ) {
        Write-PSFMessage -level host -message "current folderID: $($txtBoxFolderID.text)"
        # Creating Filter variables
        $StartDate = $FromDatePicker.Value
        $EndDate = $ToDatePicker.Value
        $MsgSubject = $txtBoxSubject.text
        $array = New-Object System.Collections.ArrayList

        $filter = "ReceivedDateTime ge $StartDate and receivedDateTime lt $EndDate"
        if ($MsgSubject -ne "") {
            $filter += " and Subject -eq '$MsgSubject'"
        }
        
        $msgs = Get-MgUserMailFolderMessage -UserId $conn.Account -MailFolderId $txtBoxFolderID.text -Filter $filter | Select-Object subject, @{N="Sender";E={$_.Sender.EmailAddress.Address}}, ReceivedDateTime, isRead
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