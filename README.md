# EXO Graph GUI Tool

## About
Graph tool to perform different operations in Exchange Online.  
This tool is the new version replacing [EwsGUI](https://github.com/agallego-css/EwsGUI).  
This tool will connect using Oauth to connect to Exchange Online. If "Modern Authentication" is not enabled in the tenant, the tool will fail to connect.  

## Pre-requisites

 > This Module requires Powershell 5.1 and above. It should work fine in PS7 and PS5.1.  
 > This Module will install different Microsoft.Graph.* modules, in order to use graph to connect to Exchange Online.  
 > Graph scopes required: "Mail.ReadWrite", "Mail.Send", "MailboxSettings.Read"  
 
## Installation

Opening Powershell with "Run as Administrator" you can run:
``` powershell
Install-Module ExoGraphGUI -Force
```
Once the module is installed, you can run:
``` powershell
Start-ExoGraphGUI
```

If you want to check for module updates you can run (the tool will already check for updates automatically):
``` powershell
Find-Module ExoGraphGUI
```
If there is any newer version than the one you already have, you can run:
``` powershell
Update-Module ExoGraphGUI -Force
```

## Authentication options

To connect to Exchange Online, it will use Modern auth and we have 2 options, either with Delegated Permission or Application permission.  
Please check on the following page for more details and options to configure your ExoGraphGUI module.
[Authentication Options](/docs/AuthenticationOptions.md)  

## Module features:
### Allows to perform 12 different operations using EWS API:
- Option 1 : List Folders in Root
- Option 2 : List folders in Recoverable Items Root folder
- Option 3 : List Items in a desired Folder
- Option 4 : Create a custom Folder in Root
- Option 5 : Delete a Folder
- Option 6 : Get user's Inbox Rules
- Option 7 : Get user's OOF Settings
- Option 8 : Move items between folders
- Option 9 : Delete a subset of items in a folder
- Option 10 : Get user's Delegate information
- Option 11 : Send Mail message
- Option 12 : Inject sample messages in the user's inbox with or without attachment
- Option 13 : Switch to another Mailbox

## Module logging

The module offers the command `Export-ExoGraphGuiLog` in order to export module logs to CSV file and/or to Powershell GridView.  
More info [here](/docs/Export-ExoGraphGuiLog.md).  

## Version History
[Change Log](/ExoGraphGUI/changelog.md)