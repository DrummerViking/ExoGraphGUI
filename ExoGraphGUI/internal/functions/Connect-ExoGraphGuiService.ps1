Function Connect-ExoGraphGuiService {
    <#
    .SYNOPSIS
    Function to authenticate the user or App.
    
    .DESCRIPTION
    Function to authenticate the user or App.
    
    .PARAMETER ClientID
    String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    String parameter with the TenantID of your AzureAD tenant.

    .PARAMETER ClientSecret
    String parameter with the Client Secret which is configured in the AzureAD App.

    .PARAMETER CertificateThumbprint


    .EXAMPLE
    PS C:\> Connect-ExoGraphGuiService
    Authenticates the user or Azure App.
    #>
    [Cmdletbinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $CertificateThumbprint,

        [String] $ClientSecret
    )
    # Connect to Graph if there is no current context
    $conn = Get-MgContext
    if ( $null -eq $conn -or $conn.Scopes -notcontains "Calendars.Read" ) {
        Write-PSFMessage -Level Host -Message "There is currently no active connection to MgGraph or current connection is missing required scopes."
        # Connecting to graph using Azure App Application flow
        if ( $clientID -ne '' -and $TenantID -ne '' -and ($CertificateThumbprint -ne '' -or $ClientSecret -ne '')) {
            Write-PSFMessage -Level Host -Message "Connecting to graph with Azure AppId: $ClientID"
            Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
        }
        else {
            # Connecting to graph with the user account
            Write-PSFMessage -Level Host -Message "Connecting to graph with the user account"
            Connect-MgGraph -Scopes "Mail.ReadBasic"
        }
        $conn = Get-MgContext
    }
    if ( $null -eq $conn.Account ) {
        Write-PSFMessage -Level Host -Message "Currently connect with App Account: $($conn.AppName)"
    }
    else {
        Write-PSFMessage -Level Host -Message "Currently connected with User Account: $($conn.Account)"
    }
    return $conn
}