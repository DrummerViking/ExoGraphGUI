Function Method12 {
    <#
    .SYNOPSIS
    Method to Delete a subset of items in a folder.
    
    .DESCRIPTION
    Method to Delete a subset of items in a folder using Date Filters and/or subject.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadWrite
    Application: Mail.ReadWrite

    .PARAMETER Account
    User's UPN to get delete messages from.
    
    .PARAMETER FolderID
    FolderID value to get mail messages from.
    
    .PARAMETER StartDate
    StartDate to search for items.
    
    .PARAMETER EndDate
    EndDate to search for items.

    .PARAMETER MsgSubject
    Optional parameter to search based on a subject text.

    .EXAMPLE
    PS C:\> Method12
    Method to Delete a subset of items in a folder.

    #>
    [CmdletBinding()]
    param(
        [String] $Account,
        [String] $FolderID,
        [string] $StartDate,
        [string] $EndDate,
        [String] $MsgSubject
    )
    $statusBarLabel.Text = "Running..."

    if ( $FolderID -ne "" )
    {
        # Creating Filter variables
        $filter = $null
        if ($MsgSubject -ne "") {
            $filter = "Subject eq '$MsgSubject'"
        }
        
        $array = New-Object System.Collections.ArrayList
        [int]$i = 0

        $msgs = Get-MgUserMailFolderMessage -UserId $Account -MailFolderId $folderID -Filter $filter -All | Where-Object { $_.ReceivedDateTime -ge $StartDate -and $_.ReceivedDateTime -lt $EndDate } | Select-Object id, subject, ReceivedDateTime
        foreach ( $msg in $msgs ) {
            $i++
            Remove-MgUserMessage -UserId $Account -MessageId $msg.Id
            $output = $msg | Select-Object @{Name="Action";Expression={"Deleting Item"}}, ReceivedDateTime, Subject
            $array.Add($output)
        }

        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBarLabel.text = "Ready. Deleted items: $i"
        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 12" -Target $Account
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}