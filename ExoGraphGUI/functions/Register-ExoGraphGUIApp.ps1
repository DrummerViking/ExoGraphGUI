Function Register-ExoGraphGUIApp {
    <#
    .SYNOPSIS
    Function to create the Azure App Registration for ExoGraphGUI.

    .DESCRIPTION
    Function to create the Azure App Registration for ExoGraphGUI.
    It will require an additional PS module "Microsoft.Graph.Applications", if not already installed it will download it.
    It will use these app scopes for the app "Mail.ReadWrite", "Mail.Send", "MailboxSettings.Read".
    You can use the "UseClientSecret" switch parameter to configure a new ClientSecret for the app. If this parameter is ommitted, we will use a Certificate.
    You can pass a certificate path if you have an existing certificate, or leave the parameter blank and a new self-signed certificate will be created.

    .PARAMETER AppName
    The friendly name of the app registration. By default will be "ExoGraphGUI Registered App".

    .PARAMETER TenantId
    Optional parameter to set the TenantID GUID.

    .PARAMETER CertPath
    The file path to your .CER public key file. If this parameter is ommitted, and the "UseClientSecret" is not used, we will be creating a new self-signed certificate (with a validity period of 1 year) for the app.

    .PARAMETER UseClientSecret
    Use this optional parameter, to configure a ClientSecret (with a validity period of 1 year) instead of a certificate.

    .PARAMETER ImportAppDataToModule
    Use this optional parameter to import your app's ClientId, TenantId and ClientSecret into the ExoGraphGUI module. In this way, the next time you run the app it will use the Application flow to authenticate with these values.

    .EXAMPLE
    PS C:\> Register-ExoGraphGUIApp.ps1 -AppName "Graph DemoApp"

    The Function will create a new AzureAD App Registration.
    The name of the app will be "ExoGraphGui Registered App".
    It will add the following API Permissions: **"Mail.ReadWrite", "MailboxSettings.Read"**.
    It will use a self-signed Certificate.

    Once the app is created, the Function will expose the link to grant "Admin consent" for the permissions requested.
    
    .NOTES
    General notes
#>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory = $false)]
        [String]
        $AppName = "ExoGraphGUI Registered App",

        [Parameter(Mandatory = $false)]
        [String]
        $TenantId,

        [Parameter(Mandatory = $false)]
        [String]
        $CertPath,

        [Parameter(Mandatory = $false)]
        [Switch]
        $UseClientSecret,
    
        [Parameter(Mandatory = $false)]
        [Switch]
        $ImportAppDataToModule
    )
    # Required modules
    Write-PSFMessage -Level Verbose -Message "Looking for required 'Microsoft.Graph.Applications' powershell module"
    if ( -not(Get-module "Microsoft.Graph.Applications" -ListAvailable) ) {
        Install-Module "Microsoft.Graph.Applications" -Scope CurrentUser -Force
    }
    Import-Module "Microsoft.Graph.Applications"

    # Graph permissions variables
    #$graphResourceId = "00000002-0000-0ff1-ce00-000000000000"
    $graphResourceId = "00000003-0000-0000-c000-000000000000"
    
    $scopesArray = New-Object System.Collections.ArrayList
    @("Mail.ReadWrite", "Mail.Send", "MailboxSettings.Read") | ForEach-Object {
        New-Variable perm -Value @{
            Id   = (Find-MgGraphPermission -SearchString $_ -PermissionType Application -ExactMatch).id
            Type = "Role"
        }
        $null = $scopesArray.add($perm)
        remove-variable perm
    }

    # Get context for access to tenant ID
    $context = Get-MgContext
    if ( $null -eq $context -or $context.Scopes -notcontains "Application.ReadWrite.All") {
        # Requires an admin
        Write-PSFMessage -Level Important -Message "Connecting to MgGraph"
        if ($TenantId) {
            Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -TenantId $TenantId
        }
        else {
            Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read"
        }
    }
    
    # Load cert
    if ( $CertPath ) {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
        Write-PSFMessage -Level Host -Message "Certificate loaded from Path '$CertPath'."
    }
    elseif ( -not($UseClientSecret) ) {
        # Create certificate
        $docsPath = [Environment]::GetFolderPath("myDocuments")
        $mycert = New-SelfSignedCertificate -DnsName $context.Account.Split("@")[1] -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(1) -KeySpec KeyExchange

        # Export certificate to .pfx file
        $mycert | Export-PfxCertificate -FilePath "$docsPath\exographgui_cert.pfx" -Password (ConvertTo-SecureString -String "LS1setup!" -AsPlainText -Force ) -Force

        # Export certificate to .cer file
        $mycert | Export-Certificate -FilePath "$docsPath\exographgui_mycert.cer" -Force
        $cerPath = Get-ChildItem -Path "$docsPath\exographgui_mycert.cer"
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath.FullName)
        Write-PSFMessage -Level Host -Message "Certificate created in Path '$($cerPath.FullName)'."
    }

    # Create app registration
    if ( -not($UseClientSecret) ) {
        $appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
            -Web @{ RedirectUris = "http://localhost"; } `
            -RequiredResourceAccess @{ ResourceAppId = $graphResourceId; ResourceAccess = $scopesArray.ToArray() } `
            -AdditionalProperties @{} -KeyCredentials @(@{ Type = "AsymmetricX509Cert"; Usage = "Verify"; Key = $cert.RawData })
        Write-PSFMessage -Level Important -Message "App registration created with app ID $($appRegistration.AppId)"
        Write-PSFMessage -Level Important -Message "You can now connect running: Start-ExoGraphGUI -ClientID $($appRegistration.AppId) -TenantID $($context.TenantId) -CertificateThumbprint $($cert.Thumbprint)"
    }
    else {
        $appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
            -Web @{ RedirectUris = "http://localhost"; } `
            -RequiredResourceAccess @{ ResourceAppId = $graphResourceId; ResourceAccess = $scopesArray.ToArray() } `
            -AdditionalProperties @{}

        $appObjId = Get-MgApplication -Filter "AppId eq '$($appRegistration.Appid)'"
        $passwordCred = @{
            displayName = 'Secret created in PowerShell'
            endDateTime = (Get-Date).Addyears(1)
        }
        $secret = Add-MgApplicationPassword -applicationId $appObjId.Id -PasswordCredential $passwordCred
        Write-PSFMessage -Level Important -Message "App registration created with app ID $($appRegistration.AppId)"
        Write-PSFMessage -Level Important -Message "Please take note of your client secret as it will not be shown anymore"
        Write-PSFMessage -Level Important -Message "You can now connect running: Start-ExoGraphGUI -ClientID $($appRegistration.AppId) -TenantID $($context.TenantId) -ClientSecret $($secret.SecretText)"
    }
    
    # Create corresponding service principal
    New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
    Write-PSFMessage -Level Verbose -Message "Service principal created"
    
    # Generate admin consent URL
    $adminConsentUrl = "https://login.microsoftonline.com/" + $context.TenantId + "/adminconsent?client_id=" + $appRegistration.AppId
    Write-PSFMessage -Level Important -Message "Please go to the following URL in your browser to provide admin consent:"
    Write-PSFMessage -Level Important -Message "$adminConsentUrl"

    if ( $ImportAppDataToModule ) {
        if ( -not($UseClientSecret) ) {
            Import-ExoGraphGUIAADAppData -ClientID $appRegistration.AppId -TenantID $context.TenantId -CertificateThumbprint $($cert.Thumbprint)
        }
        else {
            Import-ExoGraphGUIAADAppData -ClientID $appRegistration.AppId -TenantID $context.TenantId -ClientSecret $secret.SecretText
        }
    }
}