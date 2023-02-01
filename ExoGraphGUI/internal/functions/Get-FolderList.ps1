Function Get-FolderList {
    <#
    .SYNOPSIS
    Method to list folders in the user mailbox.
    
    .DESCRIPTION
    Method to list folders in the user mailbox, showing Folder name, FolderId, Number of items, and number of subfolders.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadBasic
    Application: Mail.ReadBasic.All
    
    .PARAMETER Account
    User's UPN to get mail folders from.

    .EXAMPLE
    PS C:\> Get-FolderList
    lists folders in the user mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "")]
    [CmdletBinding()]
    param(
        $Account
    )
    $statusBarLabel.Text = "Running..."

    Function Find-Subfolders {
        Param (
            $Account,

            $array,

            $ParentFolderId,

            $ParentDisplayname
        )
        foreach ($folder in (Get-MgUserMailFolderChildFolder -UserId $Account -MailFolderId $ParentFolderId -All -Property *)) {
            $folderpath = $ParentDisplayname + $folder.DisplayName
            $line = $folder | Select-Object @{N="FolderPath";E={$folderpath}},ChildFolderCount,TotalItemCount,UnreadItemCount,Id
            $null = $array.add($line)
            if ( $folder.ChildFolderCount -gt 0 ) {
                Find-Subfolders -Account $Account -ParentFolderId $folder.id -Array $array -ParentDisplayname "$folderpath\"
            }
        }
    }

    #listing all available folders in the mailbox
    $array = New-Object System.Collections.ArrayList
    if ($radiobutton1.Checked) {
        $parentFolders = (Get-MgUserMailFolder -UserId $Account -MailFolderId "msgfolderRoot").Id
    }
    elseif ($radiobutton2.Checked) {
        $deletions = Get-MgUserMailFolder -UserId $Account -MailFolderId "recoverableitemsdeletions"
        $parentFolders = $deletions.ParentFolderId
    }
    Find-Subfolders -Account $Account -ParentFolderId $parentFolders -Array $array -ParentDisplayname "\"
    
    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.Text = "Ready. Folders found: $($array.Count)"
    if ($radiobutton1.Checked) {
        Write-PSFMessage -Level Output -Message "Succesfully listed folders in the primary Mailbox" -FunctionName "Method 1" -Target $Account
    }
    elseif ($radiobutton2.Checked) {
        Write-PSFMessage -Level Output -Message "Succesfully listed folders in Recoverable Items" -FunctionName "Method 2" -Target $Account
    }
}