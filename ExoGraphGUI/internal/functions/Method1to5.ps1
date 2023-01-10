Function Method1to5 {
    <#
    .SYNOPSIS
    Method to list folders in the user mailbox.
    
    .DESCRIPTION
    Method to list folders in the user mailbox, showing Folder name, FolderId, Number of items, and number of subfolders.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadBasic
    Application: Mail.ReadBasic.All
    
    .EXAMPLE
    PS C:\> Method1to5
    lists folders in the user mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "")]
    [CmdletBinding()]
    param(
        # Parameters
    )
    $statusBarLabel.Text = "Running..."

    Function Find-Subfolders {
        Param (
            $array,

            $ParentFolderId,

            $ParentDisplayname
        )
        foreach ($folder in (Get-MgUserMailFolderChildFolder -UserId $conn.Account -MailFolderId $ParentFolderId -All -Property *)) {
            $folderpath = $ParentDisplayname + $folder.DisplayName
            $line = $folder | Select-Object @{N="FolderPath";E={$folderpath}},ChildFolderCount,TotalItemCount,UnreadItemCount,Id
            $line
            $null = $array.add($line)
            if ( $folder.ChildFolderCount -gt 0 ) {
                Find-Subfolders -ParentFolderId $folder.id -Array $array -ParentDisplayname "$folderpath\"
            }
        }
    }

    #listing all available folders in the mailbox
    $array = New-Object System.Collections.ArrayList
    if ($radiobutton1.Checked) {
        $parentFolders = (Get-MgUserMailFolder -UserId $conn.Account -MailFolderId "msgfolderroot").Id
    }
    elseif ($radiobutton4.Checked) {
        $deletions = Get-MgUserMailFolder -UserId $conn.Account -MailFolderId "recoverableitemsdeletions"
        $parentFolders = $deletions.ParentFolderId
    }
    Find-Subfolders -ParentFolderId $parentFolders -Array $array -ParentDisplayname "\"
    
    $dgResults.datasource = $array
    $dgResults.AutoResizeColumns()
    $dgResults.Visible = $True
    $txtBoxResults.Visible = $False
    $PremiseForm.refresh()
    $statusBarLabel.Text = "Ready. Folders found: $($array.Count)"
    Write-PSFMessage -Level Output -Message "Task finished succesfully" -FunctionName "Method 1-5" -Target $email
}