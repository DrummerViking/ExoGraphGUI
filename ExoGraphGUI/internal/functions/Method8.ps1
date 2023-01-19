Function Method8 {
    <#
    .SYNOPSIS
    Method to move items between folders.
    
    .DESCRIPTION
    Method to move items between folders by using FolderID values.
    Module required: Microsoft.Graph.Users.Actions
    Scope needed:
    Delegated: Mail.ReadWrite
    Application: Mail.ReadWrite

    .PARAMETER Account
    User's UPN to get move messages from.
    
    .PARAMETER FolderID
    FolderID value to get mail messages from.

    .PARAMETER TargetFolderID
    FolderID value to move mail messages to.
    
    .PARAMETER StartDate
    StartDate to search for items.
    
    .PARAMETER EndDate
    EndDate to search for items.

    .PARAMETER MsgSubject
    Optional parameter to search based on a subject text.
    
    .EXAMPLE
    PS C:\> Method8

    Moves items from source folder to target folder based on dates and/or subject filters.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "")]
    [CmdletBinding()]
    Param(
        [String] $Account,
        [String] $FolderID,
        [String] $TargetFolderID,
        [string] $StartDate,
        [string] $EndDate,
        [String] $MsgSubject
    )
    $statusBarLabel.Text = "Running..."

    if ( $FolderID -ne "" -and $TargetFolderID -ne "") {
        # Creating Filter variables
        $filter = $null
        if ($MsgSubject -ne "") {
            $filter = "Subject eq '$MsgSubject'"
        }
        
        $array = New-Object System.Collections.ArrayList
        $params = @{
            DestinationId = $TargetFolderID
        }
        $msgs = Get-MgUserMailFolderMessage -UserId $Account -MailFolderId $folderID -Filter $filter -All | Where-Object { $_.ReceivedDateTime -ge $StartDate -and $_.ReceivedDateTime -lt $EndDate } | Select-Object id, subject, @{N = "Sender"; E = { $_.Sender.EmailAddress.Address } }, ReceivedDateTime, isRead
        
        [int]$i = 0
        foreach ( $msg in $msgs ) {
            $i++
            $output = $msg | Select-Object @{Name = "Action"; Expression = { "Moving Item" } }, ReceivedDateTime, Subject
            Move-MgUserMessage -UserId $Account -MessageId $msg.Id -BodyParameter $params
            $array.Add($output)
        }
        $dgResults.datasource = $array
        $dgResults.AutoResizeColumns()
        $dgResults.Visible = $True
        $txtBoxResults.Visible = $False
        $PremiseForm.refresh()
        $statusBarLabel.text = "Ready. Moved Items: $i"
        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 8" -Target $Account
    }
    else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox or TargetFolderID is empty. Check and try again", [Microsoft.VisualBasic.MsgBoxStyle]::Okonly, "Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}