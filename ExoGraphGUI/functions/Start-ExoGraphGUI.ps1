Function Start-ExoGraphGUI {
    <#
    .SYNOPSIS
        Allows to perform 11 different operations using Graph with Exchange Online.

    .DESCRIPTION
        Allows to perform 11 different operations using Graph with Exchange Online:
        1) List Folders in Root
        2) List folders in Recoverable Items Root folder
        3) List Items in a desired Folder
        4) Create a custom Folder in Root
        5) Delete a Folder
        6) Get user's Inbox Rules
        7) Get user's OOF Settings
        8) Move items between folders
        9) Delete a subset of items in a folder
        10) Get user's Delegate information
        11) Send mail message
        12) Switch to another Mailbox
    
    .PARAMETER ClientID
    This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

    .PARAMETER TenantID
    This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

    .PARAMETER CertificateThumbprint
    This is an optional parameter. String parameter with the certificate thumbprint which is configured in the AzureAD App.

    .PARAMETER ClientSecret
    This is an optional parameter. String parameter with the Client Secret which is configured in the AzureAD App.

    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.

    .EXAMPLE
    PS C:\ Start-ExoGraphGUI
    Runs the GUI tool to use with Exchange Online.

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSReviewUnusedParameter", "")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param(
        [String] $ClientID,

        [String] $TenantID,

        #[Parameter(Mandatory = $false, ParameterSetName="Certificate")]
        [String] $CertificateThumbprint,

        #[Parameter(Mandatory = $false, ParameterSetName="ClientSecret")]
        [String] $ClientSecret
    )
    $script:nl = "`r`n"
    #$ProgressPreference = "SilentlyContinue"

    $runspaceData = Start-ModuleUpdate -ModuleRoot $script:ModuleRoot
    function GenerateForm {
         
        #region Import the Assemblies
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName Microsoft.VisualBasic
        [System.Windows.Forms.Application]::EnableVisualStyles()
        #endregion
     
        #region Generated Form Objects
        $PremiseForm = New-Object System.Windows.Forms.Form
        $radiobutton1 = New-Object System.Windows.Forms.RadioButton
        $radiobutton2 = New-Object System.Windows.Forms.RadioButton
        $radiobutton3 = New-Object System.Windows.Forms.RadioButton
        $radiobutton4 = New-Object System.Windows.Forms.RadioButton
        $radiobutton5 = New-Object System.Windows.Forms.RadioButton
        $radiobutton6 = New-Object System.Windows.Forms.RadioButton
        $radiobutton7 = New-Object System.Windows.Forms.RadioButton
        $radiobutton8 = New-Object System.Windows.Forms.RadioButton
        $radiobutton9 = New-Object System.Windows.Forms.RadioButton
        $radiobutton10 = New-Object System.Windows.Forms.RadioButton
        $radiobutton11 = New-Object System.Windows.Forms.RadioButton
        $radiobutton12 = New-Object System.Windows.Forms.RadioButton
        $radiobutton13 = New-Object System.Windows.Forms.RadioButton
        $radiobutton14 = New-Object System.Windows.Forms.RadioButton
        $radiobutton15 = New-Object System.Windows.Forms.RadioButton
        $radiobutton16 = New-Object System.Windows.Forms.RadioButton
        $labImpersonation = New-Object System.Windows.Forms.Label
        $buttonGo = New-Object System.Windows.Forms.Button
        $buttonExit = New-Object System.Windows.Forms.Button

        $dgResults = New-Object System.Windows.Forms.DataGridView
        $txtBoxResults = New-Object System.Windows.Forms.Label
        $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
        #endregion Generated Form Objects

        # Connecting to EWS and creating service object
        $service = Connect-ExoGraphGuiService -ClientID $ClientID -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint
        $Global:Account = $service.Account
        if (-not($service.Account)) {
            $Global:Account = [Microsoft.VisualBasic.Interaction]::InputBox("Enter user's UPN to work with", "ExoGraphGUI", "")
        }

        $ExpandFilters = {
            # Removing all controls, in order to reload the screen appropiately for each selection
            $PremiseForm.Controls.RemoveByKey("FromDate")
            $PremiseForm.Controls.RemoveByKey("FromDatePicker")
            $PremiseForm.Controls.RemoveByKey("ToDate")
            $PremiseForm.Controls.RemoveByKey("ToDatePicker")
            $PremiseForm.Controls.RemoveByKey("labSubject")
            $PremiseForm.Controls.RemoveByKey("txtBoxSubject")
            $PremiseForm.Controls.RemoveByKey("labFolderID")
            $PremiseForm.Controls.RemoveByKey("txtBoxFolderID")
            $PremiseForm.Controls.RemoveByKey("txtBoxTargetFolderID")
            $PremiseForm.Controls.RemoveByKey("labTargetFolderID")
            $PremiseForm.Controls.RemoveByKey("labToRecipients")
            $PremiseForm.Controls.RemoveByKey("txtBoxToRecipients")
            $PremiseForm.Controls.RemoveByKey("labCCRecipients")
            $PremiseForm.Controls.RemoveByKey("txtBoxCCRecipients")
            $PremiseForm.Controls.RemoveByKey("labBCCRecipients")
            $PremiseForm.Controls.RemoveByKey("txtBoxBCCRecipients")
            $PremiseForm.Controls.RemoveByKey("labMailSubject")
            $PremiseForm.Controls.RemoveByKey("txtBoxMailSubject")
            $PremiseForm.Controls.RemoveByKey("labMailbody")
            $PremiseForm.Controls.RemoveByKey("txtboxMailBody")
            #$PremiseForm.Controls.RemoveByKey("labelCombobox")
            #$PremiseForm.Controls.RemoveByKey("comboBoxMenu")
            #$PremiseForm.Controls.RemoveByKey("labelComboboxFolder")
            #$PremiseForm.Controls.RemoveByKey("comboBoxFolder")
            #$PremiseForm.Controls.RemoveByKey("labelComboboxConfig")
            #$PremiseForm.Controls.RemoveByKey("comboBoxConfig")

            $labFromDate = New-Object System.Windows.Forms.Label
            $global:FromDatePicker = New-Object System.Windows.Forms.DateTimePicker
            $labToDate = New-Object System.Windows.Forms.Label
            $global:ToDatePicker = New-Object System.Windows.Forms.DateTimePicker
            $labSubject = New-Object System.Windows.Forms.Label
            $global:txtBoxSubject = New-Object System.Windows.Forms.TextBox
            $labFolderID = New-Object System.Windows.Forms.Label
            $global:txtBoxFolderID = New-Object System.Windows.Forms.TextBox
            $labTargetFolderID = New-Object System.Windows.Forms.Label
            $global:txtBoxTargetFolderID = New-Object System.Windows.Forms.TextBox
            #$labelCombobox = New-Object System.Windows.Forms.Label
            #$global:comboBoxMenu = New-Object System.Windows.Forms.ComboBox
            #$labelComboboxFolder = New-Object System.Windows.Forms.Label
            #$global:comboBoxFolder = New-Object System.Windows.Forms.ComboBox
            #$labelComboboxConfig = New-Object System.Windows.Forms.Label
            #$global:comboBoxConfig = New-Object System.Windows.Forms.ComboBox
            $labToRecipients = New-Object System.Windows.Forms.Label
            $global:txtBoxToRecipients = New-Object System.Windows.Forms.TextBox
            $labCCRecipients = New-Object System.Windows.Forms.Label
            $global:txtBoxCCRecipients = New-Object System.Windows.Forms.TextBox
            $labBCCRecipients = New-Object System.Windows.Forms.Label
            $global:txtBoxBCCRecipients = New-Object System.Windows.Forms.TextBox
            $labMailSubject = New-Object System.Windows.Forms.Label
            $global:txtboxMailSubject = New-Object System.Windows.Forms.TextBox
            $labMailBody = New-Object System.Windows.Forms.Label
            $global:txtBoxMailBody = New-Object System.Windows.Forms.TextBox

            #Label TO recipients
            $labToRecipients.Location = New-Object System.Drawing.Point(5, 287)
            $labToRecipients.Size = New-Object System.Drawing.Size(25, 25)
            $labToRecipients.Name = "labToRecipients"
            $labToRecipients.Text = "To:"

            #TextBox TO recipients
            $txtBoxToRecipients.Location = New-Object System.Drawing.Point(75, 285)
            $txtBoxToRecipients.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxToRecipients.Name = "txtBoxToRecipients"
            $txtBoxToRecipients.Text = ""

            #Label CC Recipients
            $labCCRecipients.Location = New-Object System.Drawing.Point(5, 312)
            $labCCRecipients.Size = New-Object System.Drawing.Size(25, 25)
            $labCCRecipients.Name = "labCCRecipients"
            $labCCRecipients.Text = "Cc:"

            #TextBox CC recipients
            $txtBoxCCRecipients.Location = New-Object System.Drawing.Point(75, 310)
            $txtBoxCCRecipients.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxCCRecipients.Name = "txtBoxCCRecipients"
            $txtBoxCCRecipients.Text = ""

            #Label BCC Recipients
            $labBCCRecipients.Location = New-Object System.Drawing.Point(5, 337)
            $labBCCRecipients.Size = New-Object System.Drawing.Size(30, 25)
            $labBCCRecipients.Name = "labBCCRecipients"
            $labBCCRecipients.Text = "Bcc:"

            #TextBox BCC recipients
            $txtBoxBCCRecipients.Location = New-Object System.Drawing.Point(75, 335)
            $txtBoxBCCRecipients.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxBCCRecipients.Name = "txtBoxBCCRecipients"
            $txtBoxBCCRecipients.Text = ""

            #Label Mail Subject
            $labMailSubject.Location = New-Object System.Drawing.Point(5, 362)
            $labMailSubject.Size = New-Object System.Drawing.Size(45, 25)
            $labMailSubject.Name = "labMailSubject"
            $labMailSubject.Text = "Subject:"

            #TextBox Mail Subject
            $txtboxMailSubject.Location = New-Object System.Drawing.Point(75, 360)
            $txtboxMailSubject.Size = New-Object System.Drawing.Size(280, 20)
            $txtboxMailSubject.Name = "txtboxMailSubject"
            $txtboxMailSubject.Text = ""

            #Label Mail Body
            $labMailBody.Location = New-Object System.Drawing.Point(5, 387)
            $labMailBody.Size = New-Object System.Drawing.Size(45, 25)
            $labMailBody.Name = "labMailBody"
            $labMailBody.Text = "Body:"

            #TextBox Mail Body
            $txtboxMailBody.Location = New-Object System.Drawing.Point(75, 385)
            $txtboxMailBody.Size = New-Object System.Drawing.Size(280, 50)
            $txtboxMailBody.Name = "txtboxMailBody"
            $txtboxMailBody.Text = ""

            #Label FromDate
            $labFromDate.Location = New-Object System.Drawing.Point(5, 285)
            $labFromDate.Size = New-Object System.Drawing.Size(80, 35)
            $labFromDate.Name = "FromDate"
            $labFromDate.Text = "From or greater than"

            # FromDate Date Picker
            $FromDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
            $FromDatePicker.Location = New-Object System.Drawing.Point(100, 285)
            $FromDatePicker.Name = "FromDatePicker"
            $FromDatePicker.Text = ""

            #Label ToDate
            $labToDate.Location = New-Object System.Drawing.Point(5, 330)
            $labToDate.Name = "ToDate"
            $labToDate.Size = New-Object System.Drawing.Size(80, 40)
            $labToDate.Text = "To or less than"

            # ToDate Date Picker
            $ToDatePicker.DataBindings.DefaultDataSourceUpdateMode = 0
            $ToDatePicker.Location = New-Object System.Drawing.Point(100, 330)
            $ToDatePicker.Name = "ToDatePicker"
            $ToDatePicker.Text = ""

            #Label Subject
            $labSubject.Location = New-Object System.Drawing.Point(5, 370)
            $labSubject.Size = New-Object System.Drawing.Size(50, 20)
            $labSubject.Name = "labSubject"
            $labSubject.Text = "Subject: "
     
            #TextBox Subject
            $txtBoxSubject.Location = New-Object System.Drawing.Point(100, 370)
            $txtBoxSubject.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxSubject.Name = "txtBoxSubject"
            $txtBoxSubject.Text = ""

            #Label FolderID
            $labFolderID.Location = New-Object System.Drawing.Point(5, 400)
            $labFolderID.Size = New-Object System.Drawing.Size(55, 20)
            $labFolderID.Name = "labFolderID"
            $labFolderID.Text = "FolderID:"

            #TextBox FolderID
            $txtBoxFolderID.Location = New-Object System.Drawing.Point(100, 400)
            $txtBoxFolderID.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxFolderID.Name = "txtBoxFolderID"
            $txtBoxFolderID.Text = ""

            #Adapting FolderID and TxtBoxFolderID based on the selection
            if ($radiobutton4.Checked -or $radiobutton5.Checked) {
                $labFolderID.Location = New-Object System.Drawing.Point(5, 285)
                $txtBoxFolderID.Location = New-Object System.Drawing.Point(100, 285)
            }
            elseif ($radiobutton8.Checked) {
                $labFolderID.Size = New-Object System.Drawing.Size(95, 20)
                $labFolderID.Text = "SourceFolderID:"
            }
            elseif ($radiobutton12.Checked) {
                $labFolderID.Location = New-Object System.Drawing.Point(5, 285)
                $labFolderID.Size = New-Object System.Drawing.Size(95, 20)
                $labFolderID.Text = "E-mail Address:"
                $txtBoxFolderID.Location = New-Object System.Drawing.Point(100, 285)
            }

            #Label Target FolderID
            $labTargetFolderID.Location = New-Object System.Drawing.Point(5, 430)
            $labTargetFolderID.Size = New-Object System.Drawing.Size(95, 20)
            $labTargetFolderID.Name = "labTargetFolderID"
            $labTargetFolderID.Text = "TargetFolderID:"

            #TextBox Target FolderID
            $txtBoxTargetFolderID.Location = New-Object System.Drawing.Point(100, 430)
            $txtBoxTargetFolderID.Size = New-Object System.Drawing.Size(280, 20)
            $txtBoxTargetFolderID.Name = "txtBoxTargetFolderID"
            $txtBoxTargetFolderID.Text = ""

            <#Label Combobox
            $labelCombobox.Location = New-Object System.Drawing.Point(400, 285)
            $labelCombobox.Size = New-Object System.Drawing.Size(80, 35)
            $labelCombobox.Name = "labelCombobox"
            if ($radiobutton14.Checked) {
                $labelCombobox.Text = "Change Option"
            }
            else {
                $labelCombobox.Text = "Delete Option"
            }

            #ComboBox Menu
            $comboBoxMenu.DataBindings.DefaultDataSourceUpdateMode = 0
            $comboBoxMenu.FormattingEnabled = $True
            $comboBoxMenu.Location = New-Object System.Drawing.Point(485, 285)
            $comboBoxMenu.Name = "comboBoxMenu"
            $comboBoxMenu.add_SelectedIndexChanged($handler_comboBoxMenu_SelectedIndexChanged)

            #Label ComboboxFolder
            $labelComboboxFolder.Location = New-Object System.Drawing.Point(5, 285)
            $labelComboboxFolder.Size = New-Object System.Drawing.Size(50, 35)
            $labelComboboxFolder.Name = "labelComboboxFolder"
            $labelComboboxFolder.Text = "Folder:"

            #ComboBoxFolder
            $comboBoxFolder.DataBindings.DefaultDataSourceUpdateMode = 0
            $comboBoxFolder.FormattingEnabled = $True
            $comboBoxFolder.Location = New-Object System.Drawing.Point(55, 285)
            $comboBoxFolder.Size = New-Object System.Drawing.Size(70, 35)
            $comboBoxFolder.Name = "comboBoxFolder"
            $comboBoxFolder.Items.Add("") | Out-Null
            $comboBoxFolder.Items.Add("Root") | Out-Null
            $comboBoxFolder.Items.Add("Calendar") | Out-Null
            $comboBoxFolder.Items.Add("Inbox") | Out-Null
            $comboBoxFolder.add_SelectedIndexChanged($handler_comboBoxFolder_SelectedIndexChanged)

            #Label ComboboxConfig
            $labelComboboxConfig.Location = New-Object System.Drawing.Point(145, 285)
            $labelComboboxConfig.Size = New-Object System.Drawing.Size(75, 35)
            $labelComboboxConfig.Name = "labelComboboxConfig"
            $labelComboboxConfig.Text = "Config Name:"

            #ComboBoxConfig
            $comboBoxConfig.DataBindings.DefaultDataSourceUpdateMode = 0
            $comboBoxConfig.FormattingEnabled = $True
            $comboBoxConfig.Location = New-Object System.Drawing.Point(225, 285)
            $ComboboxConfig.Size = New-Object System.Drawing.Size(180, 35)
            $comboBoxConfig.Name = "comboBoxConfig"
            $comboBoxConfig.Items.Add("") | Out-Null
            $comboBoxConfig.Items.Add("Aggregated.OwaUserConfiguration") | Out-Null
            $comboBoxConfig.Items.Add("UserConfigurationProperties.All") | Out-Null
            $comboBoxConfig.Items.Add("OWA.AttachmentDataProvider") | Out-Null
            $comboBoxConfig.Items.Add("OWA.AutocompleteCache") | Out-Null
            $comboBoxConfig.Items.Add("OWA.SessionInformation") | Out-Null
            $comboBoxConfig.Items.Add("OWA.UserOptions") | Out-Null
            $comboBoxConfig.Items.Add("OWA.ViewStateConfiguration") | Out-Null
            $comboBoxConfig.Items.Add("Suite.Storage") | Out-Null
            $comboBoxConfig.Items.Add("UM.E14.PersonalAutoAttendants") | Out-Null
            $comboBoxConfig.Items.Add("CleanFinders") | Out-Null
            $comboBoxConfig.add_SelectedIndexChanged($handler_comboBoxConfig_SelectedIndexChanged)
            #>
            if ($radiobutton3.Checked) {
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
                $PremiseForm.Controls.Add($labFromDate)
                $PremiseForm.Controls.Add($FromDatePicker)
                $PremiseForm.Controls.Add($labToDate)
                $PremiseForm.Controls.Add($ToDatePicker)
                $PremiseForm.Controls.Add($labSubject)
                $PremiseForm.Controls.Add($txtBoxSubject)
            }
            elseif ($radiobutton4.Checked) {
                $labFolderID.Size = New-Object System.Drawing.Size(95, 20)
                $labFolderID.Text = "Folder Name:"
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
            }
            elseif ($radiobutton5.Checked) {
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
            }
            elseif ($radiobutton8.Checked) {
                $PremiseForm.Controls.Add($labFromDate)
                $PremiseForm.Controls.Add($FromDatePicker)
                $PremiseForm.Controls.Add($labToDate)
                $PremiseForm.Controls.Add($ToDatePicker)
                $PremiseForm.Controls.Add($labSubject)
                $PremiseForm.Controls.Add($txtBoxSubject)
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
                $PremiseForm.Controls.Add($labTargetFolderID)
                $PremiseForm.Controls.Add($txtBoxTargetFolderID)
            }
            elseif ($radiobutton9.Checked) {
                $PremiseForm.Controls.Add($labFromDate)
                $PremiseForm.Controls.Add($FromDatePicker)
                $PremiseForm.Controls.Add($labToDate)
                $PremiseForm.Controls.Add($ToDatePicker)
                $PremiseForm.Controls.Add($labSubject)
                $PremiseForm.Controls.Add($txtBoxSubject)
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
            }
            elseif ( $radiobutton11.Checked) {
                $PremiseForm.Controls.Add($labToRecipients)
                $PremiseForm.Controls.Add($txtBoxToRecipients)
                $PremiseForm.Controls.Add($labCCRecipients)
                $PremiseForm.Controls.Add($txtBoxCCRecipients)
                $PremiseForm.Controls.Add($labBCCRecipients)
                $PremiseForm.Controls.Add($txtBoxBCCRecipients)
                $PremiseForm.Controls.Add($labMailSubject)
                $PremiseForm.Controls.Add($txtboxMailSubject)
                $PremiseForm.Controls.Add($labMailBody)
                $PremiseForm.Controls.Add($txtboxMailBody)
            }
            # elseif ($radiobutton14.Checked) {
            #     $comboBoxMenu.Items.Add("") | Out-Null
            #     $comboBoxMenu.Items.Add("Normal") | Out-Null
            #     $comboBoxMenu.Items.Add("Personal") | Out-Null
            #     $comboBoxMenu.Items.Add("Private") | Out-Null
            #     $comboBoxMenu.Items.Add("Confidential") | Out-Null
            #     $comboBoxMenu.SelectedItem = "Normal"
            #     $PremiseForm.Controls.Add($labFromDate)
            #     $PremiseForm.Controls.Add($FromDatePicker)
            #     $PremiseForm.Controls.Add($labToDate)
            #     $PremiseForm.Controls.Add($ToDatePicker)
            #     $PremiseForm.Controls.Add($labSubject)
            #     $PremiseForm.Controls.Add($txtBoxSubject)
            #     $PremiseForm.Controls.Add($labFolderID)
            #     $PremiseForm.Controls.Add($txtBoxFolderID)
            #     $PremiseForm.Controls.Add($labelCombobox)
            #     $PremiseForm.Controls.Add($comboBoxMenu)
            # }
            # elseif ($radiobutton15.Checked) {
            #     $PremiseForm.Controls.Add($labelComboboxFolder)
            #     $PremiseForm.Controls.Add($comboBoxFolder)
            #     $PremiseForm.Controls.Add($labelComboboxConfig)
            #     $PremiseForm.Controls.Add($comboBoxConfig)
            # }
            elseif ($radiobutton12.Checked) {
                $PremiseForm.Controls.Add($labFolderID)
                $PremiseForm.Controls.Add($txtBoxFolderID)
            }
            $PremiseForm.refresh()
        }

        # $handler_comboBoxMenu_SelectedIndexChanged = {
        #     # Get the Event ID when item is selected
        #     $ComboOption = $comboBoxMenu.selectedItem.ToString()
        # }

        # $handler_comboBoxFolder_SelectedIndexChanged = {
        #     # Get the Event ID when item is selected
        #     $ComboOption1 = $comboBoxFolder.selectedItem.ToString()
        # }

        # $handler_comboBoxConfig_SelectedIndexChanged = {
        #     # Get the Event ID when item is selected
        #     $ComboOption2 = $comboBoxConfig.selectedItem.ToString()
        # }

        $OnLoadMainWindow_StateCorrection = { #Correct the initial state of the form to prevent the .Net maximized form issue
            $PremiseForm.WindowState = $InitialFormWindowState
        }

        #----------------------------------------------
        #region Generated Form Code

        $PremiseForm.Controls.Add($radiobutton1)
        $PremiseForm.Controls.Add($radiobutton2)
        $PremiseForm.Controls.Add($radiobutton3)
        $PremiseForm.Controls.Add($radiobutton4)
        $PremiseForm.Controls.Add($radiobutton5)
        $PremiseForm.Controls.Add($radiobutton6)
        $PremiseForm.Controls.Add($radiobutton7)
        $PremiseForm.Controls.Add($radiobutton8)
        $PremiseForm.Controls.Add($radiobutton9)
        $PremiseForm.Controls.Add($radiobutton10)
        $PremiseForm.Controls.Add($radiobutton11)
        if ( $null -eq $service.Account ) {
            $PremiseForm.Controls.Add($radiobutton12)
        }
        #$PremiseForm.Controls.Add($radiobutton13)
        #$PremiseForm.Controls.Add($radiobutton14)
        #$PremiseForm.Controls.Add($radiobutton15)
        #$PremiseForm.Controls.Add($radiobutton16)
        
        $PremiseForm.Controls.Add($buttonGo)
        $PremiseForm.Controls.Add($buttonExit)

        $statusBar = New-Object System.Windows.Forms.StatusStrip
        $statusBar.Name = "statusBar"
        $statusBarLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
        $null = $statusBar.Items.Add($statusBarLabel)
        $statusBarLabel.Text = "Ready..."
        $PremiseForm.Controls.Add($statusBar)
        $PremiseForm.ClientSize = New-Object System.Drawing.Size(850, 720)
        $PremiseForm.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $PremiseForm.Name = "form1"
        $PremiseForm.Text = "Managing user: $Account. Choose your Option"
        $PremiseForm.StartPosition = "CenterScreen"
        $PremiseForm.KeyPreview = $True
        $PremiseForm.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $PremiseForm.Close() } })
        #
        # radiobutton1
        #
        $radiobutton1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton1.Location = New-Object System.Drawing.Point(20, 20)
        $radiobutton1.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton1.Text = "1 - List Folders in Root"
        $radioButton1.Checked = $true
        $radiobutton1.UseVisualStyleBackColor = $True
        $radiobutton1.Add_Click({ & $ExpandFilters })
        #
        # radiobutton2
        #
        $radiobutton2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton2.Location = New-Object System.Drawing.Point(20, 50)
        $radiobutton2.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton2.Text = "2 - List folders in Recoverable Items Root folder"
        $radioButton2.Checked = $false
        $radiobutton2.UseVisualStyleBackColor = $True
        $radiobutton2.Add_Click({ & $ExpandFilters })
        #
        # radiobutton3
        #
        $radiobutton3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton3.Location = New-Object System.Drawing.Point(20, 80)
        $radiobutton3.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton3.Text = "3 - List Items in a desired Folder"
        $radiobutton3.Checked = $false
        $radiobutton3.UseVisualStyleBackColor = $True
        $radiobutton3.Add_Click({ & $ExpandFilters })
        #
        # radiobutton4
        #
        $radiobutton4.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton4.Location = New-Object System.Drawing.Point(20, 110)
        $radiobutton4.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton4.Text = "4 - Create a custom Folder in Root"
        $radiobutton4.Checked = $false
        $radiobutton4.UseVisualStyleBackColor = $True
        $radiobutton4.Add_Click({ & $ExpandFilters })
        #
        # radiobutton5
        #
        $radiobutton5.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton5.Location = New-Object System.Drawing.Point(20, 140)
        $radiobutton5.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton5.Text = "5 - Delete a Folder"
        $radiobutton5.Checked = $false
        $radiobutton5.UseVisualStyleBackColor = $True
        $radiobutton5.Add_Click({ & $ExpandFilters })
        #
        # radiobutton6
        #
        $radiobutton6.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton6.Location = New-Object System.Drawing.Point(20, 170)
        $radiobutton6.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton6.Text = "6 - Get user's Inbox Rules"
        $radiobutton6.Checked = $false
        $radiobutton6.UseVisualStyleBackColor = $True
        $radiobutton6.Add_Click({ & $ExpandFilters })
        #
        # radiobutton7
        #
        $radiobutton7.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton7.Location = New-Object System.Drawing.Point(20, 200)
        $radiobutton7.Name = "radiobutton7"
        $radiobutton7.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton7.Text = "7 - Get user's OOF Settings"
        $radiobutton7.Checked = $false
        $radiobutton7.UseVisualStyleBackColor = $True
        $radiobutton7.Add_Click({ & $ExpandFilters })
        #
        # radiobutton8
        #
        $radiobutton8.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton8.Location = New-Object System.Drawing.Point(20, 230)
        $radiobutton8.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton8.Text = "8 - Move items between folders"
        $radiobutton8.Checked = $false
        $radiobutton8.UseVisualStyleBackColor = $True
        $radiobutton8.Add_Click({ & $ExpandFilters })
        #
        # radiobutton9
        #
        $radiobutton9.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton9.Location = New-Object System.Drawing.Point(20, 260)
        $radiobutton9.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton9.TabIndex = 9
        $radiobutton9.Text = "9 - Delete a subset of items in a folder"
        $radiobutton9.Checked = $false
        $radiobutton9.UseVisualStyleBackColor = $True
        $radiobutton9.Add_Click({ & $ExpandFilters })
        #
        # radiobutton10
        #
        $radiobutton10.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton10.Location = New-Object System.Drawing.Point(400, 20)
        $radiobutton10.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton10.Text = "10 - Get user's Delegate information"
        $radiobutton10.Checked = $false
        $radiobutton10.UseVisualStyleBackColor = $True
        $radiobutton10.Add_Click({ & $ExpandFilters })
        #
        # radiobutton11
        #
        $radiobutton11.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton11.Location = New-Object System.Drawing.Point(400, 50)
        $radiobutton11.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton11.Text = "11 - Send mail message"
        $radiobutton11.Checked = $false
        $radiobutton11.UseVisualStyleBackColor = $True
        $radiobutton11.Add_Click({ & $ExpandFilters })
        #
        # radiobutton12
        #
        $radiobutton12.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton12.Location = New-Object System.Drawing.Point(400, 80)
        $radiobutton12.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton12.Text = "12 - Switch to another Mailbox"
        $radiobutton12.Checked = $false
        $radiobutton12.UseVisualStyleBackColor = $True
        $radiobutton12.Add_Click({ & $ExpandFilters })
        #
        # radiobutton13
        #
        $radiobutton13.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton13.Location = New-Object System.Drawing.Point(400, 110)
        $radiobutton13.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton13.Text = "13"
        $radiobutton13.Checked = $false
        $radiobutton13.UseVisualStyleBackColor = $True
        $radiobutton13.Add_Click({ & $ExpandFilters })
        #
        # radiobutton14
        #
        $radiobutton14.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton14.Location = New-Object System.Drawing.Point(400, 140)
        $radiobutton14.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton14.Text = "14"
        $radiobutton14.Checked = $false
        $radiobutton14.UseVisualStyleBackColor = $True
        $radiobutton14.Add_Click({ & $ExpandFilters })
        #
        # radiobutton15
        #
        $radiobutton15.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton15.Location = New-Object System.Drawing.Point(400, 170)
        $radiobutton15.Size = New-Object System.Drawing.Size(300, 15)
        $radiobutton15.Text = "15"
        $radiobutton15.Checked = $false
        $radiobutton15.UseVisualStyleBackColor = $True
        $radiobutton15.Add_Click({ & $ExpandFilters })
        #
        # radiobutton16
        #
        $radiobutton16.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
        $radiobutton16.Location = New-Object System.Drawing.Point(400, 200)
        $radiobutton16.Size = New-Object System.Drawing.Size(190, 15)
        $radiobutton16.Text = "16"
        $radiobutton16.Checked = $false
        $radiobutton16.UseVisualStyleBackColor = $True
        $radiobutton16.Add_Click({ & $ExpandFilters })

        #"Go" button
        $buttonGo.DataBindings.DefaultDataSourceUpdateMode = 0
        $buttonGo.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
        $buttonGo.Location = New-Object System.Drawing.Point(700, 20)
        $buttonGo.Size = New-Object System.Drawing.Size(50, 25)
        $buttonGo.Name = "Go"
        $buttonGo.Text = "Go"
        $buttonGo.UseVisualStyleBackColor = $True
        $buttonGo.add_Click({
                if ($radiobutton1.Checked) { Method1 -Account $Account}
                elseif ($radiobutton2.Checked) { Method1 -Account $Account}
                elseif ($radiobutton3.Checked) { Method3 -Account $Account -FolderId $txtBoxFolderID.Text -StartDate $FromDatePicker.Value.ToString("yyyy-MM-dd") -EndDate $ToDatePicker.Value.ToString("yyyy-MM-dd") -MsgSubject $txtBoxSubject.Text }
                elseif ($radiobutton4.Checked) { Method4 -Account $Account -DisplayName $txtBoxFolderID.Text }
                elseif ($radiobutton5.Checked) { Method5 -Account $Account -Folderid $txtBoxFolderID.Text }
                elseif ($radiobutton6.Checked) { Method6 -Account $Account }
                elseif ($radiobutton7.Checked) { Method7 -Account $Account }
                elseif ($radiobutton8.Checked) { Method8 -Account $Account -FolderId $txtBoxFolderID.Text -TargetFolderID $txtBoxTargetFolderID.Text -StartDate $FromDatePicker.Value.ToString("yyyy-MM-dd") -EndDate $ToDatePicker.Value.ToString("yyyy-MM-dd") -MsgSubject $txtBoxSubject.Text }
                elseif ($radiobutton9.Checked) { Method9 -Account $Account -FolderId $txtBoxFolderID.Text -StartDate $FromDatePicker.Value.ToString("yyyy-MM-dd") -EndDate $ToDatePicker.Value.ToString("yyyy-MM-dd") -MsgSubject $txtBoxSubject.Text }
                elseif ($radiobutton10.Checked) { Method10 -Account $Account }
                elseif ($radiobutton11.Checked) { Method11 -Account $Account -ToRecipients $txtBoxToRecipients.Text -CCRecipients $txtBoxCCRecipients.Text -BCCRecipients $txtBoxBCCRecipients.text -Subject $txtboxMailSubject.Text -Body $txtboxMailBody.Text }
                elseif ($radiobutton12.Checked) { $Global:Account = Method12 -Account $txtBoxFolderID.Text }
            })

        #"Exit" button
        $buttonExit.DataBindings.DefaultDataSourceUpdateMode = 0
        $buttonExit.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
        $buttonExit.Location = New-Object System.Drawing.Point(700, 50)
        $buttonExit.Size = New-Object System.Drawing.Size(50, 25)
        $buttonExit.Name = "Exit"
        $buttonExit.Text = "Exit"
        $buttonExit.UseVisualStyleBackColor = $True
        $buttonExit.add_Click({ $PremiseForm.Close(); return })
        
        #TextBox results
        $txtBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
        $txtBoxResults.Location = New-Object System.Drawing.Point(5, 460)
        $txtBoxResults.Size = New-Object System.Drawing.Size(840, 240)
        $txtBoxResults.Name = "TextResults"
        $txtBoxResults.BackColor = [System.Drawing.Color]::White
        $txtBoxResults.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
        $Font = New-Object System.Drawing.Font("Consolas", 8)
        $txtBoxResults.Font = $Font
        $PremiseForm.Controls.Add($txtBoxResults)

        #dataGrid

        $dgResults.Anchor = 15
        $dgResults.DataBindings.DefaultDataSourceUpdateMode = 0
        $dgResults.DataMember = ""
        $dgResults.Location = New-Object System.Drawing.Point(5, 460)
        $dgResults.Size = New-Object System.Drawing.Size(840, 240)
        $dgResults.Name = "dgResults"
        $dgResults.ReadOnly = $True
        $dgResults.RowHeadersVisible = $False
        $dgResults.Visible = $False
        $dgResults.AllowUserToOrderColumns = $True
        $dgResults.AllowUserToResizeColumns = $True
        $PremiseForm.Controls.Add($dgResults)

        #endregion Generated Form Code

        # Show Form
        #Save the initial state of the form
        $InitialFormWindowState = $PremiseForm.WindowState
        #Init the OnLoad event to correct the initial state of the form
        $PremiseForm.add_Load($OnLoadMainWindow_StateCorrection)
        $PremiseForm.Add_Shown({ $PremiseForm.Activate() })
        $PremiseForm.ShowDialog() | Out-Null
    } #End Function

    #Call the Function
    try {
        GenerateForm
    }
    finally {
        Stop-ModuleUpdate -RunspaceData $runspaceData
    }
}