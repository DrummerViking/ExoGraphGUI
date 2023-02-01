function Get-UserProfilePicture {
    <#
    .SYNOPSIS
    Method to get user's profile picture from graph.
    
    .DESCRIPTION
    Long description
    Module required: Microsoft.Graph.Mail
    Scope needed:
    Delegated: User.ReadWrite.All
    Application: User.ReadWrite.All

    .PARAMETER Account
    User's UPN to switch to.
    
    .EXAMPLE
    PS C:\> Get-UserProfilePicture -Account "user@domain.com"

    Gets user's profile picture from user "user@domain.com".
    #>
    [CmdletBinding()]
    param (
        [String] $Account
    )
    $statusBarLabel.Text = "Running..."

    try {
        $filepath = "$env:temp\profilephoto$(Get-Random).jpg"
        Write-PSFMessage -Level Verbose -Message "Setting profile picture downloaded file path to: $filepath" -FunctionName "Method 13" -Target $Account
        Invoke-MgGraphRequest -Method get -Uri "https://graph.microsoft.com/v1.0/users/$Account/photo/`$value" -OutputFilePath $filepath -ErrorAction Stop
        $Image = [System.Drawing.Image]::Fromfile($filepath)
        $pictureBox.Image = $Image.GetThumbnailImage(140, 140, $null, 0)
        $PremiseForm.refresh()
        $statusBarLabel.text = "Ready. Profile picture retrieved."
        Write-PSFMessage -Level Host -Message "Succesfully retrieved profile picture." -FunctionName "Method 13" -Target $Account
    }
    catch {
        Write-PSFMessage -Level Error -Message "The user doesn't seem to have a photo." -Target $Account
        $statusBarLabel.text = "Ready. The user doesn't seem to have a photo."
    }
}