#todo
# traceroute textfield afer traceroute with destination reached or max hops reached 
#after spoof mac, reset $adapterArray
# -check if connection via SSL/vpn/open
# documenteren
# -while telnetten van start de stop knop maken
# -als het scherm vergroot/ verkleint wordt, rescalen of niet sizable screen
# bij opstarten checken of mac adres is spoofed (edit bit)
# difference ' en "
# get-netadapter -name "*" -Physical => gets the real mac adres, not the spoofed
# retrieve the spoofed mac
# range test en trace route - disable start button until domain is niet leeg
# zelfde diverer label vaker toevoeggen
#
# select adapter (setselected 0 vervangen door de bekabelde verbinding)
# quiet mode when change mac, (no comfirmacion must be asked in dosbox)
# wifi driver doesnt support mac address spoofing, disable or make a workaround

####################################################################################
#.Synopsis 
#   Multiple tools for troubleshooting network issues with pinpads
#
#.Description 
#   This tool can be used to test of all needed ports are open, preform a 
#   traceroute or spoof mac address
#
#.Notes 
#  Author: M vd Heide  
# Version: 1.0
# Updated: 13.06.2019
####################################################################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

## global data
$routeringArray = ('POIP', 'Open Internet','Open Internet met SSL', 'Alle routeringen') # the last must be all
$hostArray = ('Equens','CCV','AWL','Xenturion','Adyen','all hosts') # the last must be all
$verboseOutput = $true
$logFile = $true
$timeoutMillieSec = 500
$ping = $ping = new-object System.Net.NetworkInformation.Ping
$defaultTracerouteHop = 6
[System.Collections.ArrayList]$csvData = Import-Csv $PSScriptRoot\pinconnections_data.csv
[System.Collections.ArrayList]$filteredCSVData = $csvData.Clone()
$adapterArray = get-netadapter -name "*" -Physical
$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

## global gui objects - tab single
$filteredProbe_RouteringListBox = New-Object System.Windows.Forms.ListBox
$filteredProbe_HostListBox = New-Object System.Windows.Forms.ListBox
$filteredProbe_ResultListView = New-Object System.Windows.Forms.ListView
## global gui objects - tab range
$rangeProbe_HostTextBox =  New-Object System.Windows.Forms.TextBox
$rangeProbe_BeginPortNumericUpDown = New-Object System.Windows.Forms.NumericUpDown 
$rangeProbe_EndPortNumericUpDown = New-Object System.Windows.Forms.NumericUpDown 
$rangeProbe_ResultRangeListView = New-Object System.Windows.Forms.ListView
## global gui objects - tab traceroute
$traceroute_Page = New-Object System.Windows.Forms.Tabpage 
$traceroute_HostTextBox = New-Object System.Windows.Forms.TextBox
$traceroute_HopNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
#$traceroute_StartButton = New-Object System.Windows.Forms.Button
$traceroute_ResultListView = New-Object System.Windows.Forms.ListView 
## global gui objects - tab MAC
$mac_AdapterComboBox = New-Object System.Windows.Forms.ComboBox
$mac_AdapterStatusValueLabel =  New-Object System.Windows.Forms.Label
$mac_AdapterDescriptionValueLabel =  New-Object System.Windows.Forms.Label
$mac_CurrentMacAddressLabel =  New-Object System.Windows.Forms.Label
$mac_CurrentMacTextBox =  New-Object System.Windows.Forms.TextBox
$mac_newMacAddressTextBox = New-Object System.Windows.Forms.TextBox
$mac_ValidMacLabel = New-Object System.Windows.Forms.Label
$mac_StartSpoofButton = New-Object System.Windows.Forms.Button

####################################################################################
#.Synopsis
#   this function starts the script. It initialise the form and the four tabs.
####################################################################################
function startScript{
    log "start pinconnection script"
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Diagnostic Pin Connections'
    $form.Size = New-Object System.Drawing.Size(800,600)
    $form.StartPosition = 'CenterScreen'

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10,10)
    $tabControl.Size = New-Object System.Drawing.Size(764,532)
    $form.Controls.Add($tabControl)

    initFilteredProbeTab
    initRangeProbeTab
    initTracerouteTab
    initMacTab

    $form.Topmost = $true
    $form.ShowDialog() | Out-Null
}

####################################################################################
#.Synopsis
#   This function initialise the gui of the  tab filteredProbe_Page. In this page 
#   the user can check only the the needed ports (or all ports). When a line in the 
#   result field is clicked the traceroute tab is opened. 
####################################################################################
function initFilteredProbeTab {
    $filteredProbe_Page = New-Object System.Windows.Forms.Tabpage 
    $filteredProbe_Page.DataBindings.DefaultDataSourceUpdateMode = 0
    $filteredProbe_Page.UseVisualStyleBackColor = $true
    $filteredProbe_Page.Text = "filter de te test poorten"
    $tabControl.Controls.Add($filteredProbe_Page)

    $filteredProbe_HelpLabel = New-Object System.Windows.Forms.Label
    $filteredProbe_HelpLabel.Location = New-Object System.Drawing.Point(8,8)
    $filteredProbe_HelpLabel.Size = New-Object System.Drawing.Size(748,68)
    #$filteredProbe_HelpLabel.AutoSize =$false
    $filteredProbe_HelpLabel.Text = "Op dit tabblad kun je filteren op routering en op verschillende hostcompanies.`r`nDoor op 'start' te klikken, worden de geselecteerde ipadressen/domains en poorten getest. De resuulten worden in onderstaand overzicht getoond. Door te dubbelklikken op een regel wordt het tabblad 'traceroute' geopend met de geselecteerd ip adres."
    $filteredProbe_Page.Controls.Add($filteredProbe_HelpLabel)

    $filteredProbe_dividerLabel = New-Object System.Windows.Forms.Label
    $filteredProbe_dividerLabel.Location = New-Object System.Drawing.Point(8,84)
    $filteredProbe_dividerLabel.AutoSize = $false
    $filteredProbe_dividerLabel.Height = 2
    $filteredProbe_dividerLabel.Width = 740
    $filteredProbe_dividerLabel.Text = ''
    $filteredProbe_dividerLabel.BorderStyle = 'Fixed3D'
    $filteredProbe_Page.Controls.Add($filteredProbe_dividerLabel)

    $filteredProbe_RouteringLabel = New-Object System.Windows.Forms.Label
    $filteredProbe_RouteringLabel.Location = New-Object System.Drawing.Point(8,102)
    $filteredProbe_RouteringLabel.Size = New-Object System.Drawing.Size(366,20)
    $filteredProbe_RouteringLabel.Text = 'Please select a routering:'
    $filteredProbe_Page.Controls.Add($filteredProbe_RouteringLabel)
    
    $filteredProbe_RouteringListBox.Location = New-Object System.Drawing.Point(8,130)
    $filteredProbe_RouteringListBox.Size = New-Object System.Drawing.Size(366,20)
    $filteredProbe_RouteringListBox.Height = 96
    for ([int]$counter = 0; $counter -le ($routeringArray.Length -1); $counter += 1){
        [void] $filteredProbe_RouteringListBox.Items.Add($routeringArray[$counter])
    }
    $filteredProbe_RouteringListBox.SetSelected($routeringArray.Length -1,$true)
    $filteredProbe_Page.Controls.Add($filteredProbe_RouteringListBox)

    $filteredProbe_HostLabel = New-Object System.Windows.Forms.Label
    $filteredProbe_HostLabel.Location = New-Object System.Drawing.Point(382,102)
    $filteredProbe_HostLabel.Size = New-Object System.Drawing.Size(366,20)
    $filteredProbe_HostLabel.Text = 'Please select a host:'
    $filteredProbe_Page.Controls.Add($filteredProbe_HostLabel)
    
    $filteredProbe_HostListBox.Location = New-Object System.Drawing.Point(382,130)
    $filteredProbe_HostListBox.Size = New-Object System.Drawing.Size(366,20)
    $filteredProbe_HostListBox.Height = 96
    for ([int]$counter = 0; $counter -le ($hostArray.Length -1); $counter += 1){
        [void] $filteredProbe_HostListBox.Items.Add($hostArray[$counter])
    }
    $filteredProbe_HostListBox.SetSelected($hostArray.Length -1,$true)
    $filteredProbe_Page.Controls.Add($filteredProbe_HostListBox)

    $filteredProbe_StartButton = New-Object System.Windows.Forms.Button
    $filteredProbe_StartButton.Location = New-Object System.Drawing.Point(344,234)
    $filteredProbe_StartButton.Size = New-Object System.Drawing.Size(75,23)
    $filteredProbe_StartButton.Text = 'start'
    $filteredProbe_StartButton.Add_Click({ filteredProbe_StartButtonClicked })
    $filteredProbe_Page.Controls.Add($filteredProbe_StartButton)
    
    
    $filteredProbe_ResultListView.Location = New-Object System.Drawing.Point(8,266)
    $filteredProbe_ResultListView.Size = New-Object System.Drawing.Size(740,230)
    $filteredProbe_ResultListView.View = [System.Windows.Forms.View]::Details
    $filteredProbe_ResultListView.Columns.Add("Remote host",200) | Out-Null
    $filteredProbe_ResultListView.Columns.Add("Poort",100) | Out-Null
    $filteredProbe_ResultListView.Columns.Add("Omschrijving",168) | Out-Null
    $filteredProbe_ResultListView.Columns.Add("Routering",168) | Out-Null
    $filteredProbe_ResultListView.Columns.Add("Status",100) | Out-Null
    $filteredProbe_ResultListView.FullRowSelect = $true
    $filteredProbe_ResultListView.add_click({filteredProbe_RowClicked})
    $filteredProbe_Page.controls.Add($filteredProbe_ResultListView)

}

####################################################################################
#.Synopsis
#   This function initialise the gui of the tab rangeProbe_Page. In this page the 
#   user can check a range of ports. When a line in the result field is clicked the
#   traceroute tab is opened. 
####################################################################################
function initRangeProbeTab{
    $rangeProbe_Page = New-Object System.Windows.Forms.Tabpage 
    $rangeProbe_Page.DataBindings.DefaultDataSourceUpdateMode = 0
    $rangeProbe_Page.UseVisualStyleBackColor = $true
    $rangeProbe_Page.Text = "test een range poorten"

    $rangeProbe_HelpLabel = New-Object System.Windows.Forms.Label
    $rangeProbe_HelpLabel.Location = New-Object System.Drawing.Point(8,8)
    $rangeProbe_HelpLabel.Size = New-Object System.Drawing.Size(748,68)
    $rangeProbe_HelpLabel.Text = "Op dit tabblad kun je een range poorten testen.`r`nDoor op 'start' te klikken, worden van de ingevoerde ipadres/domain en aangeven poorten getest. De resulten worden in onderstaand overzicht getoond. Door te dubbelklikken op een regel wordt het tabblad 'traceroute' geopend met de geselecteerd ip adres."
    $rangeProbe_Page.Controls.Add($rangeProbe_HelpLabel)

    $rangeProbe_dividerLabel = New-Object System.Windows.Forms.Label
    $rangeProbe_dividerLabel.Location = New-Object System.Drawing.Point(8,84)
    $rangeProbe_dividerLabel.AutoSize = $false
    $rangeProbe_dividerLabel.Height = 2
    $rangeProbe_dividerLabel.Width = 740
    $rangeProbe_dividerLabel.Text = ''
    $rangeProbe_dividerLabel.BorderStyle = 'Fixed3D'
    $rangeProbe_Page.Controls.Add($rangeProbe_dividerLabel)
    
    $rangeProbe_HostLabel =  New-Object System.Windows.Forms.Label
    $rangeProbe_HostLabel.Location = New-Object System.Drawing.Point(8,102)
    $rangeProbe_HostLabel.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_HostLabel.Text = 'de remote host'
    $rangeProbe_Page.Controls.Add($rangeProbe_HostLabel)

    $rangeProbe_HostTextBox.Location = New-Object System.Drawing.Point(382,102)
    $rangeProbe_HostTextBox.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_Page.Controls.Add($rangeProbe_HostTextBox)

    $rangeProbe_BeginPortLabel =  New-Object System.Windows.Forms.Label
    $rangeProbe_BeginPortLabel.Location = New-Object System.Drawing.Point(8,130)
    $rangeProbe_BeginPortLabel.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_BeginPortLabel.Text = 'begin port van de range' 
    $rangeProbe_Page.Controls.Add($rangeProbe_BeginPortLabel)

    $rangeProbe_BeginPortNumericUpDown.Location = New-Object System.Drawing.Point(382,130)
    $rangeProbe_BeginPortNumericUpDown.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_BeginPortNumericUpDown.Minimum = 0
    $rangeProbe_BeginPortNumericUpDown.Maximum = 65535
    $rangeProbe_Page.Controls.Add($rangeProbe_BeginPortNumericUpDown)

    $rangeProbe_EndPortLabel =  New-Object System.Windows.Forms.Label
    $rangeProbe_EndPortLabel.Location = New-Object System.Drawing.Point(8,158)
    $rangeProbe_EndPortLabel.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_EndPortLabel.Text = 'laatste port van de range' 
    $rangeProbe_Page.Controls.Add($rangeProbe_EndPortLabel)

    $rangeProbe_EndPortNumericUpDown.Location = New-Object System.Drawing.Point(382,158)
    $rangeProbe_EndPortNumericUpDown.Size = New-Object System.Drawing.Size(366,20)
    $rangeProbe_EndPortNumericUpDown.Minimum = 0
    $rangeProbe_EndPortNumericUpDown.Maximum = 65535
    $rangeProbe_Page.Controls.Add($rangeProbe_EndPortNumericUpDown)

    $rangeProbe_StartButton = New-Object System.Windows.Forms.Button
    $rangeProbe_StartButton.Location = New-Object System.Drawing.Point(344,234)
    $rangeProbe_StartButton.Size = New-Object System.Drawing.Size(75,23)
    $rangeProbe_StartButton.Text = 'start'
    $rangeProbe_StartButton.Add_Click({ rangeProbe_StartButtonClicked })
    $rangeProbe_Page.Controls.Add($rangeProbe_StartButton)

    $rangeProbe_ResultRangeListView.Location = New-Object System.Drawing.Point(8,266)
    $rangeProbe_ResultRangeListView.Size = New-Object System.Drawing.Size(740,230)
    $rangeProbe_ResultRangeListView.View = [System.Windows.Forms.View]::Details
    $rangeProbe_ResultRangeListView.Columns.Add("Remote host",200) | Out-Null
    $rangeProbe_ResultRangeListView.Columns.Add("Poort",100) | Out-Null
    $rangeProbe_ResultRangeListView.Columns.Add("Status",168) | Out-Null
    $rangeProbe_Page.controls.Add($rangeProbe_ResultRangeListView)

    $tabControl.Controls.Add($rangeProbe_Page)
}

####################################################################################
#.Synopsis
#   This function initialise the gui of the traceroute_Page. The user can perform a 
#   traceroute.
####################################################################################
function initTracerouteTab {
    $traceroute_Page.DataBindings.DefaultDataSourceUpdateMode = 0
    $traceroute_Page.UseVisualStyleBackColor = $true
    $traceroute_Page.Text = "trace route"

    $traceroute_HelpLabel = New-Object System.Windows.Forms.Label
    $traceroute_HelpLabel.Location = New-Object System.Drawing.Point(8,8)
    $traceroute_HelpLabel.Size = New-Object System.Drawing.Size(748,68)
    $traceroute_HelpLabel.Text = "Op dit tabblad kun je automagisch de pinverbinding in een keer werkend krijgen.`r`nHelaas, gewoon een traceroute, verder geen spannende dingen.."
    $traceroute_Page.Controls.Add($traceroute_HelpLabel)

    $traceroute_dividerLabel = New-Object System.Windows.Forms.Label
    $traceroute_dividerLabel.Location = New-Object System.Drawing.Point(8,84)
    $traceroute_dividerLabel.AutoSize = $false
    $traceroute_dividerLabel.Height = 2
    $traceroute_dividerLabel.Width = 740
    $traceroute_dividerLabel.Text = ''
    $traceroute_dividerLabel.BorderStyle = 'Fixed3D'
    $traceroute_Page.Controls.Add($traceroute_dividerLabel)

    $traceroute_HostLabel =  New-Object System.Windows.Forms.Label
    $traceroute_HostLabel.Location = New-Object System.Drawing.Point(8,102)
    $traceroute_HostLabel.Size = New-Object System.Drawing.Size(366,20)
    $traceroute_HostLabel.Text = 'de remote host'
    $traceroute_Page.Controls.Add($traceroute_HostLabel)

    $traceroute_HostTextBox.Location = New-Object System.Drawing.Point(382,102)
    $traceroute_HostTextBox.Size = New-Object System.Drawing.Size(366,20)
    $traceroute_Page.Controls.Add($traceroute_HostTextBox)

    $traceroute_HopLabel =  New-Object System.Windows.Forms.Label
    $traceroute_HopLabel.Location = New-Object System.Drawing.Point(8,130)
    $traceroute_HopLabel.Size = New-Object System.Drawing.Size(366,20)
    $traceroute_HopLabel.Text = 'maxmiaal aantal hops (max TTL)' 
    $traceroute_Page.Controls.Add($traceroute_HopLabel)

    $traceroute_HopNumericUpDown.Location = New-Object System.Drawing.Point(382,130)
    $traceroute_HopNumericUpDown.Size = New-Object System.Drawing.Size(366,20)
    $traceroute_HopNumericUpDown.Minimum = 0
    $traceroute_HopNumericUpDown.Maximum = 50
    $traceroute_HopNumericUpDown.Text = $defaultTracerouteHop
    $traceroute_Page.Controls.Add($traceroute_HopNumericUpDown)

    $traceroute_StartButton = New-Object System.Windows.Forms.Button
    $traceroute_StartButton.Location = New-Object System.Drawing.Point(344,234)
    $traceroute_StartButton.Size = New-Object System.Drawing.Size(75,23)
    $traceroute_StartButton.Text = 'start'
    $traceroute_StartButton.Add_Click({ traceroute_StartButtonClicked })
    $traceroute_Page.Controls.Add($traceroute_StartButton)

    $traceroute_ResultListView.Location = New-Object System.Drawing.Point(8,266)
    $traceroute_ResultListView.Size = New-Object System.Drawing.Size(740,230)
    $traceroute_ResultListView.View = [System.Windows.Forms.View]::Details
    $traceroute_ResultListView.Columns.Add("Hop",100) | Out-Null
    $traceroute_ResultListView.Columns.Add("Remote host",200) | Out-Null
    #$traceroute_ResultListView.Columns.Add("latency",100) | Out-Null
    $traceroute_ResultListView.Columns.Add("DNS name",436) | Out-Null
    $traceroute_Page.controls.Add($traceroute_ResultListView)

    $tabControl.Controls.Add($traceroute_Page)
}

####################################################################################
#.Synopsis
#   This function initialise the  gui of the mac_Page. The user can see and spoof 
#   the mac address of a selected physical network adapter.
####################################################################################
function initMacTab {
    $mac_Page = New-Object System.Windows.Forms.Tabpage 
    $mac_Page.DataBindings.DefaultDataSourceUpdateMode = 0
    $mac_Page.UseVisualStyleBackColor = $true
    $mac_Page.Text = "MAC Spoof"

    $mac_HelpLabel = New-Object System.Windows.Forms.Label
    $mac_HelpLabel.Location = New-Object System.Drawing.Point(8,8)
    $mac_HelpLabel.Size = New-Object System.Drawing.Size(748,68)
    $mac_HelpLabel.Text = "Op dit tabblad kun je mac adres spoofen van de fysieke netwerk adapters.`r`nAlleen als er een geldig MAC adres is ingevoerd en dit script met admin rechten is gestart kan het mac adres gespoofd worden."
    $mac_Page.Controls.Add($mac_HelpLabel)

    $mac_dividerLabel = New-Object System.Windows.Forms.Label
    $mac_dividerLabel.Location = New-Object System.Drawing.Point(8,84)
    $mac_dividerLabel.AutoSize = $false
    $mac_dividerLabel.Height = 2
    $mac_dividerLabel.Width = 740
    $mac_dividerLabel.Text = ''
    $mac_dividerLabel.BorderStyle = 'Fixed3D'
    $mac_Page.Controls.Add($mac_dividerLabel)

    $mac_AdapterLabel =  New-Object System.Windows.Forms.Label
    $mac_AdapterLabel.Location = New-Object System.Drawing.Point(8,102)
    $mac_AdapterLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_AdapterLabel.Text = 'Selecteer de netwerk adapter:'
    $mac_Page.Controls.Add($mac_AdapterLabel)
    
    $mac_AdapterComboBox.Location = New-Object System.Drawing.Point(382,102)
    $mac_AdapterComboBox.Size = New-Object System.Drawing.Size(366,20)
    foreach($adapter in $adapterArray){
        $mac_AdapterComboBox.Items.Add($adapter.name)| Out-Null
    }
    $mac_AdapterComboBox.DropDownStyle = "DropDownList"
    $mac_AdapterComboBox.SelectedIndex = 0
    $mac_AdapterComboBox.add_selectedIndexChanged({ mac_AdapterComboBoxItemChanged })
    $mac_Page.Controls.Add($mac_AdapterComboBox)

    $mac_AdapterDescriptionLabel =  New-Object System.Windows.Forms.Label
    $mac_AdapterDescriptionLabel.Location = New-Object System.Drawing.Point(8,130)
    $mac_AdapterDescriptionLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_AdapterDescriptionLabel.Text = 'geselecteerde interface'
    $mac_Page.Controls.Add($mac_AdapterDescriptionLabel)
   
    $mac_AdapterDescriptionValueLabel.Location = New-Object System.Drawing.Point(382,130)
    $mac_AdapterDescriptionValueLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_AdapterDescriptionValueLabel.Text = $adapterArray[$mac_AdapterComboBox.SelectedIndex].InterfaceDescription
    $mac_Page.Controls.Add($mac_AdapterDescriptionValueLabel)

    $mac_AdapterStatusLabel =  New-Object System.Windows.Forms.Label
    $mac_AdapterStatusLabel.Location = New-Object System.Drawing.Point(8,158)
    $mac_AdapterStatusLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_AdapterStatusLabel.Text = 'status' 
    $mac_Page.Controls.Add($mac_AdapterStatusLabel)

    $mac_AdapterStatusValueLabel.Location = New-Object System.Drawing.Point(382,158)
    $mac_AdapterStatusValueLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_AdapterStatusValueLabel.Text = $adapterArray[$mac_AdapterComboBox.SelectedIndex].status
    $mac_Page.Controls.Add($mac_AdapterStatusValueLabel)
   
    $mac_CurrentMacAddressLabel.Location = New-Object System.Drawing.Point(8,186)
    $mac_CurrentMacAddressLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_CurrentMacAddressLabel.Text = 'huidig MAC Adres' 
    $mac_Page.Controls.Add($mac_CurrentMacAddressLabel)
    
    $mac_CurrentMacTextBox.Location = New-Object System.Drawing.Point(382,186)
    $mac_CurrentMacTextBox.Size = New-Object System.Drawing.Size(366,20)
    $mac_CurrentMacTextBox.Text = $adapterArray[$mac_AdapterComboBox.SelectedIndex].MacAddress
    $mac_CurrentMacTextBox.ReadOnly = $true
    $mac_Page.Controls.Add($mac_CurrentMacTextBox)

    $mac_newMacAddressLabel =  New-Object System.Windows.Forms.Label
    $mac_newMacAddressLabel.Location = New-Object System.Drawing.Point(8,214)
    $mac_newMacAddressLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_newMacAddressLabel.Text = "nieuw MAC adres in (divider= -/:/none)."
    $mac_Page.Controls.Add($mac_newMacAddressLabel)
        
    $mac_newMacAddressTextBox.Location = New-Object System.Drawing.Point(382,214)
    $mac_newMacAddressTextBox.Size = New-Object System.Drawing.Size(366,20)
    $mac_newMacAddressTextBox.add_TextChanged({vaildMac})
    $mac_Page.Controls.Add($mac_newMacAddressTextBox)

    $mac_ValidMacLabel.Location = New-Object System.Drawing.Point(382,242)
    $mac_ValidMacLabel.Size = New-Object System.Drawing.Size(366,20)
    $mac_ValidMacLabel.Text = "Geen geldig MAC adres"
    $mac_Page.Controls.Add($mac_ValidMacLabel) 
    
    $mac_StartSpoofButton.Location = New-Object System.Drawing.Point(382,270)
    $mac_StartSpoofButton.Size = New-Object System.Drawing.Size(200,23)
    $mac_StartSpoofButton.Text = 'spoof MAC adres'
    $mac_StartSpoofButton.Enabled = $false
    $mac_StartSpoofButton.Add_Click({ mac_StartSpoofButtonClicked })
    $mac_Page.Controls.Add($mac_StartSpoofButton)

    $tabControl.Controls.Add($mac_Page)
}

####################################################################################
#.Synopsis
#   This function is called when the user press the filteredProbe_StartButton and
#   starts the filtering and port checking. It uses a clone of the the data from
#   the csv file, in this way the user can restart a filtering by pressing again 
#   on this button. 
####################################################################################
function filteredProbe_StartButtonClicked(){
    $filteredProbe_ResultListView.Items.clear()
    $filteredCSVData = $csvData.Clone()
    filteredProbe_filterArray ($filteredProbe_RouteringListBox.SelectedIndex +1) ($filteredProbe_HostListBox.SelectedIndex +1)
    filteredProbe_ProbeArray
}

####################################################################################
#.Synopsis
#   This function is called when the user press on a line in the result. the 
#   traceroute tab is opened and the host is paste in the host textbox.
####################################################################################
function filteredProbe_RowClicked(){
    $tabControl.SelectedTab = $traceroute_Page
    $traceroute_HostTextBox.Text = $filteredProbe_ResultListView.SelectedItems[0].SubItems[0].Text
}

####################################################################################
#.Synopsis
#   This function is called when the user press the rangeProbe_StartButton. It
#   starts the port checking and displays the results. 
####################################################################################
function rangeProbe_StartButtonClicked(){
    log "start range test"
    $rangeProbe_ResultRangeListView.Items.clear()
    $rangedHost = $rangeProbe_HostTextBox.Text
    $firstPort = [int]$rangeProbe_BeginPortNumericUpDown.Text 
    $lastPort = [int]$rangeProbe_EndPortNumericUpDown.Text 
    $client = New-Object System.Net.Sockets.TcpClient
    $requestCallback = $state = $null
    for ($portCounter = $firstPort;$portCounter -le $lastPort;$portCounter+=1){
        $open = checkPort $rangedHost $portCounter
        #display the results
        if($open){
            $portStatus = "open"
            $textColor ="Green"
        }else{
            $portStatus = "failed"
            $textColor ="Red"
        }
        log "$rangedHost op $portCounter : $portStatus" $textColor
        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($rangedHost)
        $ListViewItem.ForeColor = $textColor
        $ListViewItem.UseItemStyleForSubItems = $false
        $ListViewItem.SubItems.Add($portCounter).ForeColor = $textColor
        $ListViewItem.Subitems.Add($portStatus).ForeColor = $textColor
        $rangeProbe_ResultRangeListView.Items.Add($ListViewItem)
        $rangeProbe_ResultRangeListView.EnsureVisible($rangeProbe_ResultRangeListView.items.Count - 1)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

####################################################################################
#.Synopsis
#   This function is called when the user press the traceroute_StartButton. It 
#   starts the traceroute.
####################################################################################
function traceroute_StartButtonClicked(){
    $traceroute_ResultListView.Items.clear()
    traceroute_traceroute $traceroute_HostTextBox.Text
}

####################################################################################
#.Synopsis
#   This function is called when the user changes the selection of the 
#   mac_AdapterComboBox. It displays the information of the selected network adapter 
####################################################################################
function mac_AdapterComboBoxItemChanged(){
    $selectedAdapter = $adapterArray[$mac_AdapterComboBox.SelectedIndex]
    $selectedAdapter 
    $mac_AdapterDescriptionValueLabel.Text = $selectedAdapter.InterfaceDescription
    $mac_AdapterStatusValueLabel.Text = $selectedAdapter.status
    $mac_CurrentMacTextBox.Text = $selectedAdapter.MacAddress
}

####################################################################################
#.Synopsis
#   This function is called when the user press the mac_StartSpoofButton. It 
#   starts spoofs the mac address.
####################################################################################
function mac_StartSpoofButtonClicked(){ 
    $newMac = $mac_newMacAddressTextBox.Text.Trim().Replace(':','').Replace('-','')
    Set-NetAdapter -Name $mac_AdapterComboBox.SelectedItem -MacAddress $newMac 
    $mac_CurrentMacAddressLabel.Text = 'MAC adres *spoofed*'
    $mac_StartSpoofButton.Text = 'reset MAC adres'
    $mac_newMacAddressTextBox.Text = ''
    $adapterArray = get-netadapter -name "*" -Physical
    $mac_CurrentMacTextBox.Text = $adapterArray[$mac_AdapterComboBox.SelectedIndex].MacAddress
}

####################################################################################
#.Synopsis
#   This function stores in an array which ports neeed to be checked. THe index of 
#   the listbox matches with the data in the csv file. If there's no match, the data
#   is put in the array itemToBeRemoved. After this the array itemToBeRemoved is 
#   substracted from the array filteredCSVData.
#.Parameter filterRoutering
#   The int index of the item in the listbox. See csv file for the used index.
#.Parameter filterHosts
#   The int index of the item in the listbox. See csv file for the used index.
####################################################################################
function filteredProbe_filterArray([int]$filterRoutering, [int]$filterHosts) {
    #remove comments fom csvData
    $RemovableCSVData = $filteredCSVData.where( {"#" -eq ($_.hostCompany).Substring(0,1)}) 
    foreach($itemToBeRemoved in $RemovableCSVData){
        $filteredCSVData.Remove($itemToBeRemoved)
    }
    # remove those with the not-selected routering from csvData
    $logText = "filtering met indexen $filterRoutering, $filterHosts : van de $($filteredCSVData.Count) worden er "
    if (!($filterRoutering.Equals($routeringArray.Count))){
        $RemovableCSVData = $filteredCSVData.where( {([int]($_.routering) -NE $filterRoutering) -and ([int]($_.routering) -ne $routeringArray.Count)}) 
        $logText += "$($RemovableCSVData.count) gefilterd (routering) en blijven er "
        foreach($itemToBeRemoved in $RemovableCSVData){
            $filteredCSVData.Remove($itemToBeRemoved)
        }
        $logText += "$($filteredCSVData.Count) over. "
    }
    # remove those with the not-selected host from csvData
    if (!($filterHosts.Equals($hostArray.Count))){
        $RemovableCSVData = $filteredCSVData.where( {([int]($_.hostCompany) -NE $filterHosts) -and ([int]($_.hostCompany) -ne $hostArray.Count)})
        $logText += "Daarna worden er $($RemovableCSVData.count) gefilterd (hosts) en blijven er "
        foreach($itemToBeRemoved in $RemovableCSVData){
            $filteredCSVData.Remove($itemToBeRemoved)
        }
    }
    $logText += "$($filteredCSVData.Count) over."
    if ($filterRoutering -eq $routeringArray.Count -and $filterHosts -eq $hostArray.Count){
        $logText = "filtering met indexen $filterRoutering, $filterHosts : er wordt niet gefilterd."
    }
    log $logText
}

####################################################################################
#.Synopsis
#   This function checks the ports based on the filtered data and displays the
#   results.
####################################################################################
function filteredProbe_ProbeArray() {
    log "start poorten check"
    $filteredCSVData | ForEach { 
        $hostCompany = $_.hostCompany
        $remoteHost = $_.remoteHost
        $port = $_.port
        $routering = $_.routering
        $hostDescription = $_.hostDescription

        #test the connection
        $open = checkport $remoteHost $port

        #log the results
        if($open){
            $portStatus = "open"
            $textColor ="Green"
        }else{
            $portStatus = "failed"
            $textColor ="Red"
        }
        log "$remoteHost op port $port ($hostDescription via $($routeringArray[$routering-1])): $portStatus" $textColor
		 
		#display the results
        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($remoteHost)
        $ListViewItem.ForeColor  = $textColor
        $ListViewItem.UseItemStyleForSubItems = $false
        $ListViewItem.SubItems.Add($port).ForeColor = $textColor
        $ListViewItem.SubItems.Add($hostDescription).ForeColor = $textColor
        $ListViewItem.SubItems.Add($routeringArray[$routering-1]).ForeColor = $textColor
        $ListViewItem.Subitems.Add($portStatus).ForeColor = $textColor
        $filteredProbe_ResultListView.Items.Add($ListViewItem)
        $filteredProbe_ResultListView.EnsureVisible($filteredProbe_ResultListView.items.Count - 1)
        [System.Windows.Forms.Application]::DoEvents()
        
    }
}

function traceroute_traceroute($tracerouteDestination){
    $MaxTTL= $traceroute_HopNumericUpDown.Text
    $Fragmentation=$false
    $VerboseOutput=$true
    $Timeout=5000

    $success = [System.Net.NetworkInformation.IPStatus]::Success

    log "Tracing to $tracerouteDestination"
    for ($i=1; $i -le $MaxTTL; $i++) {
        $reply = getTracerouteHop $i $tracerouteDestination $Timeout
        $addr = $reply.Address

        try {
            $dns = [System.Net.Dns]::GetHostByAddress($addr)
        }
        catch {
            $dns = "-"
        }

        $name = $dns.HostName

        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($i)
        $ListViewItem.SubItems.Add([String]$addr)
        #$ListViewItem.SubItems.Add([String]$reply.RoundTripTime)
        $ListViewItem.SubItems.Add([String]$name)
        $traceroute_ResultListView.Items.Add($ListViewItem)
        $traceroute_ResultListView.EnsureVisible($traceroute_ResultListView.items.Count - 1)
        [System.Windows.Forms.Application]::DoEvents()
        
        #$obj | Add-Member -MemberType NoteProperty -Name latency -Value $reply.RoundTripTime

        log "Hop: $i`t= $addr`t($name)"

        if($reply.Status -eq $success){break}
    }
}

function checkPort($rh, $po){      
$ErrorActionPreference = 'SilentlyContinue'
    $client = New-Object System.Net.Sockets.TcpClient
    $requestCallback = $state = $null
        $beginConnect = $client.BeginConnect($rh,$po,$requestCallback,$state)
        Start-Sleep -milli 60
        if ($client.Connected) {
            $open = $true
        } else { 
            Start-Sleep -milli ($timeoutMillieSec - 60)
            $open = $client.Connected 
        }
        $client.Close()
        $ErrorActionPreference = 'Continue'
        
        return $open
}        

function vaildMac(){
    $regex = "((\d|([a-f]|[A-F])){2}){6}"
    $valid = $false
    $mac = $mac_newMacAddressTextBox.Text.Trim().Replace(':','').Replace('-','')
    if ($mac.Length -eq 12){
        if ($mac -match $regex){
          $valid = $true
        }
    }
    $mac_StartSpoofButton.Enabled = $valid -and $isAdmin
    if (!$valid){
        if ($isAdmin){
            $mac_ValidMacLabel.Text = "Geen geldig MAC adres"
        } else {
            $mac_ValidMacLabel.Text = "Geen geldig MAC adres en geen Admin rights"
        }
    } else {
        if ($isAdmin){
            $mac_ValidMacLabel.Text = ""
        } else {
            $mac_ValidMacLabel.Text = "Geldig MAC adres, maar geen Admin rights"
        }
    }
}


function getTracerouteHop($hop, $tracerouteDestination, $Timeout){
    $popt = new-object System.Net.NetworkInformation.PingOptions($hop, $false)   
    $reply = $ping.Send($tracerouteDestination, $Timeout, [System.Text.Encoding]::Default.GetBytes("MESSAGE"), $popt)
    return $reply
}


function log($message, $textColor){
    if ($textColor -eq $null){
        $textColor = "White"
    }
    $timestamp = Get-Date
    $timeStampedMessage = "$($timestamp.ToUniversalTime()) - $message"

    if ($logfile){
        Add-Content  $PSScriptRoot\pinconnections.log -Value $timeStampedMessage
    }
    if ($verboseOutput){
        Write-Host $timeStampedMessage -ForegroundColor $textColor
    }
}

# this command starts the script
startScript
