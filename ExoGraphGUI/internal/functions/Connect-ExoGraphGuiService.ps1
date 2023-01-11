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
    
    .PARAMETER CertificateThumbprint
    String parameter with the certificate thumbprint which is configured in the AzureAD App.

    .EXAMPLE
    PS C:\> Connect-ExoGraphGuiService
    Authenticates the user or Azure App.
    #>
    [Cmdletbinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $CertificateThumbprint
    )

    DynamicParam {
        $modules = Get-Module Microsoft.Graph.Authentication
        $latest = $modules | Sort-Object version -Descending -Top 1
        if ($latest.Version -ge [version]"2.0.0") {
            $Attribute = New-Object System.Management.Automation.ParameterAttribute
            $Attribute.Mandatory = $false
            $Attribute.HelpMessage = "String parameter with the Client Secret which is configured in the AzureAD App."

            #create an attributecollection object for the attribute we just created.
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]

            #add our custom attribute
            $attributeCollection.Add($Attribute)

            #add our paramater specifying the attribute collection
            $secretParam = New-Object System.Management.Automation.RuntimeDefinedParameter('ClientSecret', [String], $attributeCollection)

            #expose the name of our parameter
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('ClientSecret', $secretParam)
            return $paramDictionary
        }
    }

    Process {
        # Connect to Graph if there is no current context
        $conn = Get-MgContext
        if ( $null -eq $conn -or $conn.Scopes -notcontains "Calendars.Read" ) {
            Write-PSFMessage -Level Host -Message "There is currently no active connection to MgGraph or current connection is missing required scopes."
            # Connecting to graph using Azure App Application flow
            if ( $clientID -ne '' -and $TenantID -ne '' -and ($CertificateThumbprint -ne '' -or $ClientSecret -ne '')) {
                Write-PSFMessage -Level Host -Message "Connecting to graph with Azure AppId: $ClientID"
                if ($PSBoundParameters.Contains('CertificateThumbprint')) {
                    Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
                }
                elseif ($PSBoundParameters.Contains('ClientSecret') ) {
                    $clientCredential = New-Object System.Net.NetworkCredential($ClientID, $ClientSecret)
                    Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $clientCredential
                }
            }
            else {
                # Connecting to graph with the user account
                Write-PSFMessage -Level Host -Message "Connecting to graph with the user account"
                Connect-MgGraph -Scopes "Mail.ReadBasic"
            }
            $conn = Get-MgContext
        }
        if ( $null -eq $conn.Account ) {
            Write-PSFMessage -Level Host -Message "Currently connected with App Account: $($conn.AppName)"
        }
        else {
            Write-PSFMessage -Level Host -Message "Currently connected with User Account: $($conn.Account)"
        }
        return $conn
    }
}