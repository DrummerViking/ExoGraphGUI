Function New-CustomFolder {
    <#
    .SYNOPSIS
    Method to create a custom folder in mailbox's Root.
    
    .DESCRIPTION
    Method to create a custom folder in mailbox's Root.
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: Mail.ReadWrite
    Application: Mail.ReadWrite
    
    .PARAMETER Account
    User's UPN to create mail folder to.

    .PARAMETER DisplayName
    DisplayName of the folder to be created.

    .EXAMPLE
    PS C:\> New-CustomFolder
    Method to create a custom folder in mailbox's Root.

    #>
    [CmdletBinding()]
    param(
        [String] $Account,
        [String] $DisplayName
    )
    
    if ( $DisplayName -ne "" )
    {
        $statusBarLabel.text = "Running..."
 
        $params = @{
            DisplayName = $DisplayName
            IsHidden = $false
        }
        New-MgUserMailFolder -UserId $Account -BodyParameter $params

        Write-PSFMessage -Level Host -Message "Succesfully created folder: $DisplayName" -FunctionName "Method 4" -Target $Account
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("FolderID textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Method 7 finished with warnings/errors"
    }
}