Function Switch-AnotherMailbox {
    <#
    .SYNOPSIS
    Method to switch to another mailbox.
    
    .DESCRIPTION
    Method to switch to another mailbox.
    
    .PARAMETER Account
    User's UPN to switch to.
    
    .EXAMPLE
    PS C:\> Switch-AnotherMailbox
    Method to switch to another mailbox.

    #>
    [OutputType([String])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [CmdletBinding()]
    param(
        [String] $Account
    )
    $statusBarLabel.Text = "Running..."

     if ( $Account -ne "" ) {
        $labImpersonation.Location = New-Object System.Drawing.Point(440,200)
        $labImpersonation.Size = New-Object System.Drawing.Size(300,20)
        $labImpersonation.Name = "labImpersonation"
        $labImpersonation.ForeColor = "Blue"
        $PremiseForm.Controls.Add($labImpersonation)
        $labImpersonation.Text = $Account
        $PremiseForm.Text = "Managing user: " + $Account + ". Choose your Option"

        Write-PSFMessage -Level Host -Message "Succesfully switched to user account: $Account" -FunctionName "Method 15" -Target $Account
        $statusBarLabel.text = "Ready..."
        $PremiseForm.Refresh()
        return $Account
    }
    else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Email Address textbox is empty. Check and try again",[Microsoft.VisualBasic.MsgBoxStyle]::Okonly,"Information Message")
        $statusBarLabel.text = "Process finished with warnings/errors"
    }
}