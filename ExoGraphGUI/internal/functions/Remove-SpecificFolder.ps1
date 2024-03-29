﻿Function Remove-SpecificFolder {
    <#
    .SYNOPSIS
    Method to delete a specific folder in the user mailbox.
    
    .DESCRIPTION
    Method to delete a specific folder in the user mailbox with 3 different deletion methods.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadWrite
    Application: Mail.ReadWrite

    .PARAMETER Account
    User's UPN to delete mail folder from.

    .PARAMETER FolderId
    FolderId of the folder to be deleted.

    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.

    .EXAMPLE
    PS C:\> Remove-SpecificFolder
    Method to delete a specific folder in the user mailbox.

    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param(
        [String] $Account,

        [String] $FolderId
    )
    $statusBarLabel.Text = "Running..."

    if ( $FolderId -ne "" )
    {
        Remove-MgUserMailFolder -UserId $Account -MailFolderId $FolderId

        Write-PSFMessage -Level Host -Message "Succesfully removed folder $folderID." -FunctionName "Method 5" -Target $Account
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}