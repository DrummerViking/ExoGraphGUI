Function Method16 {
    <#
    .SYNOPSIS
    Method to switch to another mailbox.
    
    .DESCRIPTION
    Method to switch to another mailbox.
    
    .PARAMETER Account
    User's UPN to switch to.
    
    .EXAMPLE
    PS C:\> Method16
    Method to switch to another mailbox.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [CmdletBinding()]
    param(
        [String] $Account
    )
    $statusBarLabel.Text = "Running..."

     if ( $Account -ne "" )
    {
        $labImpersonation.Location = New-Object System.Drawing.Point(595,200)
        $labImpersonation.Size = New-Object System.Drawing.Size(300,20)
        $labImpersonation.Name = "labImpersonation"
        $labImpersonation.ForeColor = "Blue"
        $PremiseForm.Controls.Add($labImpersonation)
        $labImpersonation.Text = $Account
        $PremiseForm.Text = "Managing user: " + $Account + ". Choose your Option"

        Write-PSFMessage -Level Host -Message "Task finished succesfully" -FunctionName "Method 16" -Target $Account
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
    }
    else
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Email Address textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}