﻿Function Connect-ExoGraphGuiService {
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
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [Cmdletbinding()]
    param(
        [String] $ClientID,

        [String] $TenantID,

        [String] $CertificateThumbprint
    )

    DynamicParam {
        $modules = Get-Module Microsoft.Graph.Authentication
        $latest = $modules | Sort-Object version -Descending
        if ($latest[0].Version -ge [version]"2.0.0") {
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
        $requiredScopes = "Mail.ReadWrite", "Mail.Send", "MailboxSettings.Read", "User.ReadWrite.All"
        if ( $conn ) {
            $compare = Compare-Object -ReferenceObject $conn.scopes -DifferenceObject $requiredScopes -IncludeEqual
        }
        if ( $null -eq $conn -or $compare.sideindicator -contains "=>" ) {
            Write-PSFMessage -Level Host -Message "There is currently no active connection to MgGraph or current connection is missing required scopes: $($requiredScopes -join ", ")" -FunctionName "ExoGraphGUI"
            if ( $clientID -ne '' -and $TenantID -ne '' -and ($CertificateThumbprint -ne '' -or $ClientSecret -ne '')) {
                # Connecting to graph using Azure App Application flow with passed parameters
                Write-PSFMessage -Level Host -Message "Connecting to graph with Azure AppId: $ClientID with passed parameters" -FunctionName "ExoGraphGUI"
                if ($PSBoundParameters.ContainsKey('CertificateThumbprint') ) {
                    Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
                }
                elseif ($PSBoundParameters.ContainsKey('ClientSecret') ) {
                    $clientCredential = New-Object System.Net.NetworkCredential($ClientID, $ClientSecret)
                    Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $clientCredential
                }
            }
            elseif (
                $null -ne (Get-PSFConfig -Module ExoGraphGUI -Name ClientID).value -and `
                $null -ne (Get-PSFConfig -Module ExoGraphGUI -Name TenantID).value -and `
                ($null -ne (Get-PSFConfig -Module ExoGraphGUI -Name ClientSecret).value -or $null -ne (Get-PSFConfig -Module ExoGraphGUI -Name CertificateThumbprint).value)
            ) {
                # Connecting to graph using Azure App Application flow saved values in the module
                Write-PSFMessage -Level Host -Message "Connecting to graph with Azure AppId: $((Get-PSFConfig -Module ExoGraphGUI -Name ClientID).value) with saved credentials in the module" -FunctionName "ExoGraphGUI"
                $cid = (Get-PSFConfig -Module ExoGraphGUI -Name ClientID).value
                $tid = (Get-PSFConfig -Module ExoGraphGUI -Name TenantID).value
                $cs = ConvertTo-SecureString -String (Get-PSFConfig -Module ExoGraphGUI -Name ClientSecret).value -AsPlainText -Force
                $ct = (Get-PSFConfig -Module ExoGraphGUI -Name CertificateThumbprint).value
                if ( $ct ) {
                    Write-PSFMessage -Level Verbose -Message "Connecting to graph with Azure AppId: $cid with saved CertificateThumbprint"
                    Connect-MgGraph -ClientId $cid -TenantId $tid -CertificateThumbprint $ct
                }
                else {
                    Write-PSFMessage -Level Verbose -Message "Connecting to graph with Azure AppId: $cid with saved ClientSecret"
                    $clientCredential = New-Object System.Net.NetworkCredential($cid, $cs)
                    Connect-MgGraph -TenantId $tid -ClientSecretCredential $clientCredential
                }
            }
            else {
                # Connecting to graph with the user account
                Write-PSFMessage -Level Host -Message "Connecting to graph with the user account" -FunctionName "ExoGraphGUI"
                Connect-MgGraph -Scopes $requiredScopes
            }
            $conn = Get-MgContext
        }
        if ( $null -eq $conn.Account ) {
            Write-PSFMessage -Level Host -Message "Currently connected with App Account: $($conn.AppName)" -FunctionName "ExoGraphGUI"
        }
        else {
            Write-PSFMessage -Level Host -Message "Currently connected with User Account: $($conn.Account)" -FunctionName "ExoGraphGUI"
        }
        return $conn
    }
}