# Authentication options

## Using Delegated Permissions  

In order to connect using Modern Authentication with Delegated permissions, we just need to have the required powershell modules.
At the time of connecting to the service, if we are not passing any ClientID, TenantID and Certificate, we assume we will use a user credential to use its own mailbox.  

## Using Application Permissions

In order to connect with an Application Permission, we just need to create the Application in Azure following the steps here: [Register application for app-only authentication](https://learn.microsoft.com/en-us/graph/tutorials/powershell-app-only?view=graph-rest-1.0&tabs=windows&tutorial-step=1).  
Or you can run the powershell function to create the app for you:  
```Powershell
Register-ExoGraphGUIApp
```
The function will create a new AzureAD App Registration.  
It will download a necessary Graph Powershell module to create the app registration.  
The name of the app will be "ExoGraphGui Registered App".  
It will add the following API Permissions: **"Mail.ReadWrite", "Mail.Send", "MailboxSettings.Read"**.  
it will use a self-signed Certificate.  

Once the app is created, it will expose the link to grant "Admin consent" for the permissions requested.  

Additionally you can run the function with the parameter "ImportAppDataToModule" like this:  
```Powershell
Register-ExoGraphGUIApp -ImportAppDataToModule
```
And the script will create the AzureAD App registration as mentioned above, and will follow the below instructions to save app data into the module automatically.  


Once you create your app with a ClientSecret, you can use this tool by running:  
```Powershell
Start-ExoGraphGui -ClientID "your app client ID" -TenantID "Your tenant ID" -CertificateThumbprint "your certificate's thumbprint"
```

## Saving your Azure App details in the EWSGui module

If you want to use Application permission flow, we have an option to save your "ClientID", "TenantID" and "ClientSecret", so you don't need to enter it every time as the example above.  
you can run:  
```Powershell
Import-ExoGraphGuiAADAppData -ClientID "your app client ID" -TenantID "Your tenant ID" [-ClientSecret "your Secret passcode"] [-CertificateThumbprint "your certificate's thumbprint"]
```

Now everytime you want to run the module, just run `Start-ExoGraphGui` and will fetch these saved details (so it will follow the Application permissions flow).  
<br>
if you need to revert this change, let's say you need to try Delegated Permission back again, you can unregister these values:  
```Powershell
Remove-ExoGraphGuiAADAppData
```