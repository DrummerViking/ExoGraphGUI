@{
	# Script module or binary module file associated with this manifest
	RootModule = 'ExoGraphGUI.psm1'
	
	# Version number of this module.
	ModuleVersion = '1.1.1'
	
	# ID used to uniquely identify this module
	GUID = '01b4d9df-0d7a-4697-abf1-23173181029a'
	
	# Author of this module
	Author = 'Agustin Gallegos [MSFT]'
	
	# Company or vendor of this module
	CompanyName = 'Microsoft'
	
	# Copyright statement for this module
	Copyright = 'Copyright (c) 2023 Agustin Gallegos'
	
	# Description of the functionality provided by this module
	Description = 'Graph GUI tool to connect to Exchange Online and perform different operations'
	
	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '5.0'
	
	# Modules that must be imported into the global environment prior to importing
	# this module
	RequiredModules = @(
		@{ ModuleName='PSFramework'; ModuleVersion='1.7.249' }
		@{ ModuleName='BurntToast'; ModuleVersion='0.8.5' }
		@{ ModuleName='Microsoft.Graph.Authentication'; ModuleVersion='1.20.0' }
		#@{ ModuleName='Microsoft.Graph.Calendar'; ModuleVersion='1.20.0' }
		@{ ModuleName='Microsoft.Graph.Mail'; ModuleVersion='1.20.0' }
		#@{ ModuleName='Microsoft.Graph.Users'; ModuleVersion='1.20.0' }
		@{ ModuleName='Microsoft.Graph.Users.Actions'; ModuleVersion='1.20.0' }
	)
	
	# Assemblies that must be loaded prior to importing this module
	# RequiredAssemblies = @('bin\ExoGraphGUI.dll')
	
	# Type files (.ps1xml) to be loaded when importing this module
	# TypesToProcess = @('xml\ExoGraphGUI.Types.ps1xml')
	
	# Format files (.ps1xml) to be loaded when importing this module
	# FormatsToProcess = @('xml\ExoGraphGUI.Format.ps1xml')
	
	# Functions to export from this module
	FunctionsToExport = @(
		'Export-ExoGraphGUILog'
		'Import-ExoGraphGUIAADAppData'
		'Register-ExoGraphGUIApp'
		'Remove-ExoGraphGUIAADAppData'
		'Start-ExoGraphGUI'
	)
	
	# Cmdlets to export from this module
	CmdletsToExport = ''
	
	# Variables to export from this module
	VariablesToExport = ''
	
	# Aliases to export from this module
	AliasesToExport = ''
	
	# List of all modules packaged with this module
	ModuleList = @()
	
	# List of all files packaged with this module
	FileList = @()
	
	# Private data to pass to the module specified in ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
	PrivateData = @{
		
		#Support for PowerShellGet galleries.
		PSData = @{
			
			# Tags applied to this module. These help with module discovery in online galleries.
			Tags = @('ExchangeOnline','MSGraph','Graph')
			
			# A URL to the license for this module.
			LicenseUri = 'https://github.com/agallego-css/ExoGraphGUI/blob/master/LICENSE'
			
			# A URL to the main website for this project.
			ProjectUri = 'https://github.com/agallego-css/ExoGraphGUI/'
			
			# A URL to an icon representing this module.
			# IconUri = ''
			
			# ReleaseNotes of this module
			# ReleaseNotes = ''
			
		} # End of PSData hashtable
		
	} # End of PrivateData hashtable
}