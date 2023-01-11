function Remove-ExoGraphGUIAADAppData {
    <#
    .SYNOPSIS
    Function to remove ClientID, TenantID and ClientSecret to the ExoGraphGUI powershell module.
    
    .DESCRIPTION
    Function to remove ClientID, TenantID and ClientSecret to the ExoGraphGUI powershell module.
    
    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.

    .EXAMPLE
    PS C:\> Remove-ExoGraphGUIAADAppData

    The script will Remove these values in the ExoGraphGUI module to be used automatically.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param (
        # Parameters
    )
    
    begin {

    }
    
    process {
        Write-PSFMessage -Level Important -Message "Removing ClientID, TenantID and ClientSecret strings from ExoGraphGUI Module."
        Unregister-PSFConfig -Module ExoGraphGUI
        remove-PSFConfig -Module ExoGraphGUI -Name clientID -Confirm:$false
        remove-PSFConfig -Module ExoGraphGUI -Name tenantID -Confirm:$false
        remove-PSFConfig -Module ExoGraphGUI -Name ClientSecret -Confirm:$false
    }
    
    end {
        
    }
}