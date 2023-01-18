# EXO Graph GUI Tool

## About
Graph tool to perform different operations in Exchange Online.  
This tool is the new version replacing [EwsGUI](https://github.com/agallego-css/EwsGUI).  
This tool will connect using Oauth to connect to Exchange Online. If "Modern Authentication" is not enabled in the tenant, the tool will fail to connect.  

## Pre-requisites

 > This Module requires Powershell 5.1 and above. It should work fine in PS7 and PS5.1.  
 > This Module will install different Microsoft.Graph.* modules, in order to use graph to connect to Exchange Online.  
 > Graph scopes required: "Mail.ReadWrite", "MailboxSettings.Read"  
 
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
### Allows to perform 16 different operations using EWS API:
- Option 1 : List Folders in Root
- Option 2 : List Folders in Archive Root
- Option 3 : List Folders in Public Folder Root
- Option 4 : List folders in Recoverable Items Root folder
- Option 5 : List folders in Recoverable Items folder in Archive
- Option 6 : List Items in a desired Folder
- Option 7 : Create a custom Folder in Root
- Option 8 : Delete a Folder
- Option 9 : Get user's Inbox Rules
- Option 10 : Get user's OOF Settings
- Option 11 : Move items between folders
- Option 12 : Delete a subset of items in a folder
- Option 13 : Get user's Delegate information
- Option 14 : Change sensitivity to items in a folder
- Option 15 : Remove OWA configurations
- Option 16 : Switch to another Mailbox

## Module logging

The module offers the command `Export-ExoGraphGuiLog` in order to export module logs to CSV file and/or to Powershell GridView.  
More info [here](/docs/Export-ExoGraphGuiLog.md).  

## Version History
[Change Log](/ExoGraphGUI/changelog.md)