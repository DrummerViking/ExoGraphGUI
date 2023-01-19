function Import-ExoGraphGUIAADAppDataa {
    <#
    .SYNOPSIS
    Function to Import ClientID, TenantID and ClientSecret to the ExoGraphGUI powershell module.

    .DESCRIPTION
    Function to Import ClientID, TenantID and ClientSecret to the ExoGraphGUI powershell module.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID your AzureAD tenant.

    .PARAMETER CertificateThumbprint
    String parameter with the certificate's thumbprint associated to your AzureAD App registration.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.
    
    .EXAMPLE
    PS C:\> Import-ExoGraphGUIAADAppDataa -ClientID "your app client ID" -TenantID "Your tenant ID" -ClientSecret "your Secret passcode"

    The script will Import these values in the ExoGraphGUI module to be used automatically.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String] $ClientID,

        [Parameter(Mandatory = $true)]
        [String] $TenantID,

        [Parameter(Mandatory = $false, ParameterSetName = "Certificate")]
        $CertificateThumbprint,

        [Parameter(Mandatory = $false, ParameterSetName = "ClientSecret")]
        [String] $ClientSecret
    )

    begin {

        if ( $ClientID -eq '' -or $TenantID -eq '' -or $CertificateThumbprint -eq '' -or $ClientSecret -eq '') {
            throw "Either ClientID, TenantID or CertificateThumbprint/ClientSecret are null or empty."
        }
    }

    process {
        Write-PSFMessage -Level Important -Message "Importing ClientID string to ExoGraphGUI Module."
        Set-PSFConfig -Module ExoGraphGUI -Name "ClientID" -Value $ClientID -Description "AppID of your Azure Registered App" -AllowDelete -PassThru | Register-PSFConfig
        
        Write-PSFMessage -Level Important -Message "Importing TenantID string to ExoGraphGUI Module."
        Set-PSFConfig -Module ExoGraphGUI -Name "TenantID" -Value $TenantID -Description "TenantID where your Azure App is registered." -AllowDelete -PassThru | Register-PSFConfig
        
        if ( $PSBoundParameters.ContainsKey("CertificateThumbprint") ) {
                Write-PSFMessage -Level Important -Message "Importing CertificateThumbprint string to ExoGraphGUI Module."
                Set-PSFConfig -Module ExoGraphGUI -Name "CertificateThumbprint" -Value $CertificateThumbprint -Description "CertificateThumbprint for your Azure App" -AllowDelete -PassThru | Register-PSFConfig
            }
        elseif ($PSBoundParameters.ContainsKey("ClientSecret") ) {
            Write-PSFMessage -Level Important -Message "Importing ClientSecret string to ExoGraphGUI Module."
            Set-PSFConfig -Module ExoGraphGUI -Name "ClientSecret" -Value $ClientSecret -Description "ClientSecret passcode for your Azure App" -AllowDelete -PassThru | Register-PSFConfig
        }
    }
}