function Select-ProfilePicture {
    <#
    .SYNOPSIS
    Opens dialog box to pick new picture.
    
    .DESCRIPTION
    Opens dialog box to pick new picture.
    
    .EXAMPLE
    PS C:\> Select-ProfilePicture
    Opens dialog box to pick new picture.
    #>
    [OutputType([String])]
    [CmdletBinding()]
    param (
        # Parameters
    )
    $statusBarLabel.Text = "Running..."
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyPictures")
    $OpenFileDialog.Filter = "Image Files(*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|All files (*.*)|*.*"
    $null = $OpenFileDialog.ShowDialog()
    Write-PSFMessage -Level Verbose -Message "Selected profile picture file path: $($OpenFileDialog.FileName)" -FunctionName "Method 14" -Target $Account
    $Image = [System.Drawing.Image]::Fromfile($OpenFileDialog.FileName)
    $pictureBox.Image = $Image.GetThumbnailImage(140, 140, $null, 0)
    $PremiseForm.refresh()
    $statusBarLabel.text = "Ready. Profile picture selected."
    Write-PSFMessage -Level Host -Message "Succesfully selected new profile picture." -FunctionName "Method 14" -Target $Account
    return $OpenFileDialog.FileName
}