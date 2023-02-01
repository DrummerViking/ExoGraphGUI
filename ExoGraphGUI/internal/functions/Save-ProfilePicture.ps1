function Save-ProfilePicture {
    <#
    .SYNOPSIS
    Saves new profile picture into the user's pofile.
    
    .DESCRIPTION
    Saves new profile picture into the user's pofile.
    Module required: Microsoft.Graph.Authentication
    Scope needed:
    Delegated: User.ReadWrite.All
    Application: User.ReadWrite.All
    
    .PARAMETER Account
    User's UPN to save the new picture to.
    
    .PARAMETER NewProfilePicture
    File path to the new profile picture.
    
    .EXAMPLE
    PS C:\> Save-ProfilePicture -Account $Account -NewProfilePicture "C:\temp\photo.jpg"

    Saves photo "C:\temp\photo.jpg" to the user $Account.
    #>
    [CmdletBinding()]
    param (
        [String] $Account,
        [String] $NewProfilePicture
    )
    try {
        $statusBarLabel.Text = "Running..."
        Write-PSFMessage -Level Verbose -Message "uploading profile picture from: $NewProfilePicture" -FunctionName "Method 14" -Target $Account
        if ( $PSVersionTable.PSVersion.Major -lt 7) {
            $requestBody = Get-Content $NewProfilePicture -Raw -Encoding Byte
        }
        else {
            $requestBody = Get-Content $NewProfilePicture -AsByteStream -Raw
        }
        Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/users/$Account/photo/`$value" -Body $requestBody -ContentType "image/jpeg" -ErrorAction Stop
        
        $statusBarLabel.text = "Ready. Profile picture saved."
        Write-PSFMessage -Level Host -Message "Succesfully saved profile picture." -FunctionName "Method 14" -Target $Account
    }
    catch {
        Write-PSFMessage -Level Error -Message "Something failed to upload new profile picture. Error message. $_" -Target $Account
        $statusBarLabel.text = "Ready. Something failed to upload new profile picture."
    }
}