$SearchSingleFindButtonClick = {
    
    $SearchSingleClearQueueButton.Enabled = $false
    $SearchSingleSendTestPrintButton.Enabled = $false
    $SearchSingleFindButton.Enabled = $false
    $SearchSingleEditPrinterQueueButton.Enabled = $false
    $SearchSingleRDPToXXXXButton.Enabled = $false
    $SearchSinglegridview.Rows.Clear()
    $result = @()
    $printername = ($SearchSingleprinterNameBox.Text).ToUpper()

    #search active directory for queue server names
    $serverlist = @()
    $root = new-object system.DirectoryServices.DirectoryEntry
    $search = new-object system.DirectoryServices.DirectorySearcher($root)
    $search.PageSize = 1000
    $search.Filter = "(&(objectCategory=printQueue) (printername=$printername))"
    $search.PropertyNamesOnly = $true
    $search.PropertiesToLoad.Add('ShortServerName')
    $search.SearchScope = "subtree"
    $searchresult = $search.FindAll()

foreach($result in $searchresult){

	$serverlist += $result.GetDirectoryEntry().ShortServerName

}

    
    foreach($servername in $serverlist){

        if($servername -ne $null){

            $array = New-Object System.Collections.ArrayList
            start-job -Name PrintJobs_$servername -ScriptBlock{
                $servername = $args[0]
                $printerName = $args[1]
                get-wmiobject win32_perfFormattedData_Spooler_PrintQueue -ComputerName $servername -Filter "Name = '$printerName'" | select name,jobs
            } -ArgumentList $servername,$Printername
            start-job -Name PrinterInfo_$servername -ScriptBlock{
                $servername = $args[0]
                $printerName = $args[1]
                [WMI]('\\' + $servername + '\root\cimv2:win32_printer.DeviceID="' + $printername + '"')
            } -ArgumentList $servername,$Printername

            get-job -Name PrintJobs_$servername, PrinterInfo_$servername | Wait-Job | Out-Null

            $script:PrintJobs = Receive-Job PrintJobs_$servername
            $script:PrinterInfo = Receive-Job PrinterInfo_$servername

            ForEach($printer in $printerInfo) {

                $job = $printjobs | Where-Object { $_.name -eq $printer.name} | select -ExpandProperty jobs
                $portvalue = [WMI]('\\' + $servername + '\root\cimv2:win32_TCPIPPrinterPort.Name="' + $printer.portname + '"') | select -ExpandProperty HostAddress                 
                $dnspingcheck = [WMI]('\\' + $servername + '\root\cimv2:win32_pingstatus.Address="' + $portvalue + '"')                
                if ($dnspingcheck.StatusCode -eq 0){$ping = 'Responding to ping'}else{$ping = 'Request timed out'}
                
                 if ($printer.drivername -like "HP*"){
                    if ($printer.printerstate -eq "128" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Offline' }
                    elseif ($printer.printerstate -eq "1154" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Error - Offline'}
                    elseif ($printer.printerstate -eq "16" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Out of Paper'}
                    elseif ($printer.printerstate -eq "131072" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Toner/Ink Low'}
                    elseif ($printer.printerstate -eq "262144" -and $printer.printerstatus -eq "1"){ $printerstatus = 'No Toner/Ink'}
                    elseif ($printer.printerstate -eq "0" -and $printer.printerstatus -eq "2"){ $printerstatus = 'Ready - Stuck Job'}
                    elseif ($printer.printerstate -eq "0" -and $printer.printerstatus -eq "3"){ $printerstatus = 'Ready -Idle'}
                    elseif ($printer.printerstatus -eq "4"){ $printerstatus = 'Printing' } else{ $printerstatus = "unknown"}

                }

                if ($printer.drivername -like "XEROX*"){
                
                    if ($printer.printerstate -eq "2" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Error' }
                    elseif ($printer.printerstate -eq "1154" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Error - Offline'}
                    elseif ($printer.printerstate -eq "16" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Out of Paper'}
                    elseif ($printer.printerstate -eq "131072" -and $printer.printerstatus -eq "1"){ $printerstatus = 'Toner/Ink Low'}
                    elseif ($printer.printerstate -eq "0" -and $printer.printerstatus -eq "2"){ $printerstatus = 'Ready - Stuck Job'}
                    elseif ($printer.printerstate -eq "0" -and $printer.printerstatus -eq "3"){ $printerstatus = 'Ready -Idle'}
                    elseif ($printer.printerstatus -eq "4"){ $printerstatus = 'Printing' } else{ $printerstatus = "unknown"}
                
                }
                                           
                $result = New-Object -TypeName PSObject -Property @{
                    'ServerName' = $servername;
                    'PrinterName' = $printer.name;
                    'JobsInQueue' = $job;
                    'PortName' = $printer.portname;
                    'PortValue' = $portvalue;
                    'DriverName' = $printer.drivername
                    'Comment' = $printer.comment;
                    'Location' = $printer.location;
                    'Status' = $printerstatus;
                    'PingCheck' = $ping;

                } #newobject end
               

            } #foreach printer in printerinfo end
             $array.AddRange(@($result))
                $SearchSinglegridview.Rows.Add($array.ServerName, $array.PrinterName,$array.Status, $array.PingCheck, $array.JobsInQueue,$array.PortName, $array.PortValue, $array.DriverName, $array.Comment, $array.Location)
                $SearchSinglegridview.autoresizecolumns()
                $SearchSinglePrinter.Refresh()

        } #if servername not null end

    } #foreach servername in serverlist end
    [system.GC]::Collect()
    $SearchSingleFindButton.Enabled = $true
    $SearchSingleClearQueueButton.Enabled = $true
    $SearchSingleSendTestPrintButton.Enabled = $true
    $SearchSingleEditPrinterQueueButton.Enabled = $false
    $SearchSingleRDPToXXXXButton.Enabled = $true

} #findbuttonclick end

$SearchSingleClearQueueClick = {

$ServerName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[0].Value }
$PrinterName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[1].Value }

$cancel =  [WMI]('\\' + $ServerName + '\root\cimv2:win32_printer.DeviceID="' + $PrinterName + '"')
$return = $cancel.CancelAllJobs() | select -expand returnvalue

If($return -eq 0){

    [System.Windows.Forms.MessageBox]::Show("Successfully cleared printer queue for $PrinterName on $ServerName")

}else{[System.Windows.Forms.MessageBox]::Show("Unable to clear printer queue for $PrinterName on $ServerName")}
      [system.GC]::Collect() 
} #SearchSingleClearQueueClick end

$SearchSingleSendTestPrintClick = {
$ServerName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[0].Value }
$PrinterName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[1].Value }

$cancel =  [WMI]('\\' + $ServerName + '\root\cimv2:win32_printer.DeviceID="' + $PrinterName + '"')
$return = $cancel.PrintTestPage() | select -expand returnvalue

If($return -eq 0){

    [System.Windows.Forms.MessageBox]::Show("Test Page sent to $PrinterName on $ServerName")

}else{[System.Windows.Forms.MessageBox]::Show("Unable to send Test Page to $PrinterName on $ServerName")}
      [system.GC]::Collect() 

} #SearchSingleSendTestPrintClick end

$SearchSingleEditPrinterQueueClick = {
    
    $EditQueueClick = {
            $EditButton.Enabled = $false
            $DriverName = $DriverNameComboBox.Text
            $PortName = $PortNameBox.Text
            $PortValue = $PortValueBox.Text
            $Comment = $CommentBox.Text
            $Location = $LocationBox.Text

            #   Printer Port Creation
            $portclass = [wmiclass]"\\$ServerName\root\cimv2:Win32_TcpIpPrinterPort"
            $newPort = $portclass.CreateInstance()
            $newport.Name= "$PortName"
            $newport.SNMPEnabled=$false 
            $newport.Protocol=1 
            $newport.HostAddress= "$PortValue" 
            $newport.Put() > $null

            #    Printer Queue Creation
            $print = "\\$ServerName\root\cimv2:Win32_Printer"
            $print = [WMICLASS]$print
            $print.scope.options.enableprivileges = $true
            $newprinter = $print.createInstance() 
            $newprinter.drivername = "$DriverName"
            $newprinter.PortName = "$PortName"
            $newprinter.Shared = $true
            $newprinter.network = $true
            $newprinter.Sharename = $PrinterName
            $newprinter.Location = "$Location"
            $newprinter.Comment = $Comment
            $newprinter.DeviceID = $PrinterName
            $newprinter.Published = $true
            $newprinter.DoCompleteFirst = $true
            $newprinter.EnableXXXDI = $true
            $newprinter.Put() > $null

            $EditButton.Enabled = $true

            [System.Windows.Forms.MessageBox]::Show("Completed!")
            [system.GC]::Collect()
            $EditQueueForm.Close()
    }
    
    $ServerName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[0].Value }
    $PrinterName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[1].Value }
    $PortName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[6].Value }
    $PortValue = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[7].Value }
    $DriverName = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[8].Value }
    $Comment = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[9].Value }
    $Location = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[10].Value }

    $DriverList = get-wmiobject -class win32_PrinterDriver -ComputerName $ServerName | select @{n="DriverName";e={(($_.name).split(",") | select -index 0)}} | select -unique -ExpandProperty DriverName
    $CurrentDriverIndex = [array]::indexof($DriverList,$DriverName)


    #EditPrinterQueueClick Build Form
    $EditQueueForm = new-object System.Windows.Forms.Form
    $EditQueueForm.Width = 520
    $EditQueueForm.Height = 320
    $EditQueueForm.Text = "Edit Print Queue"
    $EditQueueForm.StartPosition = "CenterScreen"
    
    #Add Controls
    $PrinterNameLabel = new-object System.Windows.Forms.Label
    $PrinterNameLabel.Location = new-object System.Drawing.Size(15,27)
    $PrinterNameLabel.Size = new-object System.Drawing.Size(73,23)
    $PrinterNameLabel.Text = "PrinterName:"
    
    $PrinterNameBox = new-object System.Windows.Forms.TextBox
    $PrinterNameBox.Location = new-object System.Drawing.Size(95,23)
    $PrinterNameBox.Size = new-object System.Drawing.Size(385,23)
    $printerNameBox.Text = $PrinterName
    $PrinterNameBox.ReadOnly = $true

    $ServerNameLabel = new-object System.Windows.Forms.Label
    $ServerNameLabel.Location = new-object System.Drawing.Size(15,57)
    $ServerNameLabel.Size = new-object System.Drawing.Size(73,23)
    $ServerNameLabel.Text = "ServerName:"
    
    $ServerNameBox = new-object System.Windows.Forms.TextBox
    $ServerNameBox.Location = new-object System.Drawing.Size(95,53)
    $ServerNameBox.Size = new-object System.Drawing.Size(385,23)
    $ServerNameBox.Text = $ServerName
    $ServerNameBox.ReadOnly = $true

    $PortNameLabel = new-object System.Windows.Forms.Label
    $PortNameLabel.Location = new-object System.Drawing.Size(15,87)
    $PortNameLabel.Size = new-object System.Drawing.Size(73,23)
    $PortNameLabel.Text = "PortName:"
    
    $PortNameBox = new-object System.Windows.Forms.TextBox
    $PortNameBox.Location = new-object System.Drawing.Size(95,83)
    $PortNameBox.Size = new-object System.Drawing.Size(385,23)
    $PortNameBox.Text = $PortName

    $PortValueLabel = new-object System.Windows.Forms.Label
    $PortValueLabel.Location = new-object System.Drawing.Size(15,117)
    $PortValueLabel.Size = new-object System.Drawing.Size(73,23)
    $PortValueLabel.Text = "PortValue:"
    
    $PortValueBox = new-object System.Windows.Forms.TextBox
    $PortValueBox.Location = new-object System.Drawing.Size(95,113)
    $PortValueBox.Size = new-object System.Drawing.Size(385,23)
    $PortValueBox.Text = $PortValue

    $DriverNameLabel = new-object System.Windows.Forms.Label
    $DriverNameLabel.Location = new-object System.Drawing.Size(15,147)
    $DriverNameLabel.Size = new-object System.Drawing.Size(73,23)
    $DriverNameLabel.Text = "DriverName:"
    
    $DriverNameComboBox = new-object System.Windows.Forms.ComboBox
    $DriverNameComboBox.Location = new-object System.Drawing.Size(95,143)
    $DriverNameComboBox.Size = new-object System.Drawing.Size(385,23)
    $DriverNameComboBox.XXXndingContext = $EditQueueForm.XXXndingContext
    $DriverNameComboBox.DataSource = $DriverList
    $DriverNameComboBox.SelectedIndex = $CurrentDriverIndex

    $CommentLabel = new-object System.Windows.Forms.Label
    $CommentLabel.Location = new-object System.Drawing.Size(15,177)
    $CommentLabel.Size = new-object System.Drawing.Size(73,23)
    $CommentLabel.Text = "Comment:"
    
    $CommentBox = new-object System.Windows.Forms.TextBox
    $CommentBox.Location = new-object System.Drawing.Size(95,173)
    $CommentBox.Size = new-object System.Drawing.Size(385,23)
    $CommentBox.Text = $Comment

    $LocationLabel = new-object System.Windows.Forms.Label
    $LocationLabel.Location = new-object System.Drawing.Size(15,207)
    $LocationLabel.Size = new-object System.Drawing.Size(73,23)
    $LocationLabel.Text = "Location:"
    
    $LocationBox = new-object System.Windows.Forms.TextBox
    $LocationBox.Location = new-object System.Drawing.Size(95,203)
    $LocationBox.Size = new-object System.Drawing.Size(385,23)
    $LocationBox.Text = $Location
 
    $EditButton = new-object System.Windows.Forms.Button
    $EditButton.Location = new-object System.Drawing.Size(220,240)
    $EditButton.Size = new-object System.Drawing.Size(70,30)
    $EditButton.Text = "Start Edit"

    $EditQueueForm.controls.Add($PrinterNameLabel)
    $EditQueueForm.controls.Add($PrinterNameBox)
    $EditQueueForm.controls.Add($ServerNameLabel)
    $EditQueueForm.controls.Add($ServerNameBox)
    $EditQueueForm.controls.Add($PortNameLabel)
    $EditQueueForm.controls.Add($PortNameBox)
    $EditQueueForm.controls.Add($PortValueLabel)
    $EditQueueForm.controls.Add($PortValueBox)
    $EditQueueForm.controls.Add($DriverNameLabel)
    $EditQueueForm.controls.Add($DriverNameComboBox)
    $EditQueueForm.controls.Add($CommentLabel)
    $EditQueueForm.controls.Add($CommentBox)
    $EditQueueForm.controls.Add($LocationLabel)
    $EditQueueForm.controls.Add($LocationBox)
    $EditQueueForm.controls.Add($EditButton)

    #Add Events
    $EditButton.Add_Click( $EditQueueClick )

    $EditQueueForm.ShowDialog()
    [system.GC]::Collect()

} #SearchSingleEditPrinterQueueClick end

$SearchSingleRDPToXXXXClick = {

    $servername = $SearchSinglegridview.SelectedRows | ForEach-Object { $_.Cells[0].Value }
    mstsc /v:$servername

}

$CreateNewPrinterSearchClick = {
    $pct = 1
    $steps = 6
    [int]$totalpct = ($pct / $steps ) * 100
    $NewServerNameComboBox.items.Clear()
    $NewDriverNameComboBox.items.Clear()
    $NewMacAddressBox.text = $null
    $NewPrinterNameBox.text = $null
    $NewPortNameBox.text = $null
    $NewPortValueBox.text = $null
    $NewMakeModelBox.text = $null
    $NewLocationBox.text = $null
    $NewPrinterProvisionPrinterButton.Enabled = $false
    $NewAssetTagBox.text = $null
    $NewSerialNoBox.text = $null
    $NewTapBox.text = $null
    $newProgressBar.Visible = $true
    $newProgressBar.Value = $totalpct
    $NewPrinterSearchButton.Enabled = $false
    $NewIPAddressBox.Enabled = $false
    $scopelist = @()
    $result = @()
    $ipaddress = $NewIPAddressBox.Text
    $ipaddress = ($ipaddress).trim()
    $pct++
    [int]$totalpct = ($pct / $steps ) * 100
    $ipcheck = $ipaddress -match ("^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$")

    if ($ipcheck -eq $false)    {
        [System.Windows.Forms.MessageBox]::Show("$ipaddress is invalid. Please check and try again")
        
    }else{
        $leasecheck = @()
        $macaddress = @()
        $scope = @()
        $sitecode = @()
        $serverlist = @()
        $serverlistfinal = @()
        $NewDriverList = @()
        $dhcpserver = @()
        $dhcpserverslist = get-content "$PSScriptRoot\DHCPServers.txt"
        
        foreach ($dhcpserverlist in $dhcpserverslist){
            $dhcpserver += Get-ciminstance win32_pingstatus -Filter "address='$dhcpserverlist'" | Where-Object {$_.StatusCode -eq 0} |  select -ExpandProperty address
        }
        
        $pct++
        [int]$totalpct = ($pct / $steps ) * 100
        $newProgressBar.Value = $totalpct

        $dhcpserver = Get-Random $dhcpserver
        
            foreach ($dhcpserverl in $dhcpserverslist){
                
                $lease = Invoke-wmiMethod -Class PS_DhcpServerv4Lease -Namespace root/microsoft/windows/dhcp -ComputerName $dhcpserverl -Argument $null,$ipaddress -Name GetByIPAddress | select -expand properties | select -expand value | select -first 1 AddressState,ClientId,ClientType,HostName,IPAddress,ScopeId,__SERVER
                                
                    $macaddress += $lease.clientid
                    $scope += $lease.scopeid
                    $scopeserver = $dhcpserverl
                    $leasecheck += $lease.addressstate -like "*Reservation*"

                    if ($lease -ne $null){
                        
                       $sitecode += Invoke-wmiMethod -Class PS_DhcpServerv4scope -Namespace root/microsoft/windows/dhcp -ComputerName $scopeserver -name get -ArgumentList $null, $lease.scopeid | select -ExpandProperty cmdletOutput | select -ExpandProperty name
                    
                    }
            
            }
  
            $scope = $scope | select -first 1

            if ($scope -ne $null -and $leasecheck -notcontains "true"){
                    $serverlist = @()
                    $macaddress = $macaddress | select -first 1
                    $sitecode = ($sitecode[0].split("-") | select -Index 1 -ErrorAction SilentlyContinue ).trim()
                    $root = new-object system.DirectoryServices.DirectoryEntry
                    $search = new-object system.DirectoryServices.DirectorySearcher($root)
                    $search.PageSize = 1000
                    $search.Filter = "(&(objectCategory=printQueue) (printername=PRXXX$sitecode*))"
                    $search.PropertyNamesOnly = $true
                    $search.PropertiesToLoad.Add('ShortServerName') | Out-Null
                    $search.PropertiesToLoad.Add('PrinterName') | Out-Null
                    $search.SearchScope = "subtree"
                    $searchresult = $search.FindAll()
                     
                    $pct++
                    [int]$totalpct = ($pct / $steps ) * 100
                    $newProgressBar.Value = $totalpct

                    foreach($result in $searchresult){
                        [System.Windows.Forms.Application]::DoEvents()
	                    $serverlist += $result.GetDirectoryEntry().ShortServerName
                        $printerlist += $result.GetDirectoryEntry().PrinterName

                    }
                    

                    if ($printerlist -eq $null){

                        $printername = "XXXXX$sitecode" + "0001"

                        $serverlist = @()
                        $root = new-object system.DirectoryServices.DirectoryEntry
                        $search = new-object system.DirectoryServices.DirectorySearcher($root)
                        $search.PageSize = 1000
                        $search.Filter = "(&(objectCategory=Computer) (name=*XXXX*))"
                        $search.PropertyNamesOnly = $true
                        $search.PropertiesToLoad.Add('Name') | Out-Null
                        $search.SearchScope = "subtree"
                        $searchresult = $search.FindAll()

                        foreach ($s in $searchresult){
                            [System.Windows.Forms.Application]::DoEvents()
                            $serverlist += $s.GetDirectoryEntry().Name | Sort-Object -Descending
                        }
                        $pct++
                        [int]$totalpct = ($pct / $steps ) * 100
                        $newProgressBar.Value = $totalpct

                        foreach ($server in $serverlist){ 
                    
                            $serverlistfinal += $server + ": " + "0" + " installed printers"  
                    
                        }
                        $servername = $serverlist[0]

                    }else{
                        $pct++
                        [int]$totalpct = ($pct / $steps ) * 100
                        $newProgressBar.Value = $totalpct
                        $lastprinter = $printerlist | Sort-Object -Descending | select -first 1
                        $printergetting= $lastprinter.substring(8)
                        $printergetting = $printergetting.trim("PS")
                        [int]$printernumber = $printergetting
                        $printernumber++
                        $printername = $lastprinter.substring(0,8) + ("{0:0000}" -f $printernumber) + "PS"
                        $alllastprinters = gc '\\UKXXXXXXSXDDC001\C$\CVS\LastPrinters.txt'
                        do {
                        $lastprintersnull = $alllastprinters | Where-Object { $_ -eq $printername }
                        if ($lastprintersnull -ne $null){

                        $printergetting= $printername.substring(8)
                        $printergetting = $printergetting.trim("PS")
                        [int]$printernumber = $printergetting
                        $printernumber++
                        $printername = $printername.substring(0,8) + ("{0:0000}" -f $printernumber) + "PS"

                        }

                        }until( $lastprintersnull -eq $null )
                    
                        $serverlist = $serverlist | Group-Object | select name,count | Sort-Object count 

                        foreach ($server in $serverlist){ 
                            [System.Windows.Forms.Application]::DoEvents()
                            $serverlistfinal += $server.name + ": " + $server.count + " installed printers"  
                    
                        }
                        $servername = $serverlist[0].name
                    }
                    $pct++
                    [int]$totalpct = ($pct / $steps ) * 100
                    $newProgressBar.Value = $totalpct
                                        
                    $NewDriverList = get-ciminstance -class win32_PrinterDriver -ComputerName $ServerName | select @{n="DriverName";e={(($_.name).split(",") | select -index 0)}} | select -unique -ExpandProperty DriverName

                    $NewMacAddressBox.text = $macaddress
                       
                    $NewServerNameComboBox.Items.addrange(@($serverlistfinal))
                    $NewServerNameComboBox.SelectedIndex = 0
                    $NewPrinterNameBox.text = $printername
                    $NewPortNameBox.text = $printername
                    $NewPortValueBox.text = $ipaddress
                    $NewDriverNameComboBox.SelectedIndex = 0
                    $NewMacAddressLabel.enabled = $true 
                    $NewMacAddressBox.enabled = $true 
                    $NewPrinterNameLabel.enabled = $true
                    $NewPrinterNameBox.enabled = $true
                    $NewServerNameLabel.enabled = $true
                    $NewServerNameComboBox.enabled = $true
                    $NewPortNameLabel.enabled = $true
                    $NewPortNameBox.enabled = $true
                    $NewPortValueLabel.enabled = $true
                    $NewPortValueBox.enabled = $true
                    $NewDriverNameLabel.enabled = $true
                    $NewDriverNameComboBox.enabled = $true
                    $NewAssetTagLabel.enabled = $true
                    $NewAssetTagBox.enabled = $true
                    $NewSerialNoLabel.enabled = $true
                    $NewSerialNoBox.enabled = $true
                    $NewTapLabel.enabled = $true
                    $NewTapBox.enabled = $true
                    $NewLocationLabel.enabled = $true
                    $NewLocationBox.enabled = $true
                    $NewMakeModelLabel.enabled = $true
                    $NewMakeModelBox.enabled = $true
                    $NewServerNameListAllCheckBox.enabled = $true
                    $NewServerNameListAllLabel.enabled = $true
                    $NewPrinterProvisionPrinterButton.Enabled = $true

                }
            elseif($leasecheck -contains "true"){

                [System.Windows.Forms.MessageBox]::Show("$ipaddress has an existing DHCP reservation. Please manually check.")
                    
                    $NewMacAddressLabel.enabled = $false 
                    $NewMacAddressBox.enabled = $false 
                    $NewPrinterNameLabel.enabled = $false
                    $NewPrinterNameBox.enabled = $false
                    $NewServerNameLabel.enabled = $false
                    $NewServerNameComboBox.enabled = $false
                    $NewPortNameLabel.enabled = $false
                    $NewPortNameBox.enabled = $false
                    $NewPortValueLabel.enabled = $false
                    $NewPortValueBox.enabled = $false
                    $NewDriverNameLabel.enabled = $false
                    $NewDriverNameComboBox.enabled = $false
                    $NewAssetTagLabel.enabled = $false
                    $NewAssetTagBox.enabled = $false
                    $NewSerialNoLabel.enabled = $false
                    $NewSerialNoBox.enabled = $false
                    $NewTapLabel.enabled = $false
                    $NewTapBox.enabled = $false
                    $NewLocationLabel.enabled = $false
                    $NewLocationBox.enabled = $false
                    $NewServerNameListAllCheckBox.enabled = $false
                    $NewServerNameListAllLabel.enabled = $false
                    $NewPrinterProvisionPrinterButton.Enabled = $false

            }else{
                    [System.Windows.Forms.MessageBox]::Show("unable to find a lease under $ipaddress")
                     
                    $NewMacAddressLabel.enabled = $false 
                    $NewMacAddressBox.enabled = $false 
                    $NewPrinterNameLabel.enabled = $false
                    $NewPrinterNameBox.enabled = $false
                    $NewServerNameLabel.enabled = $false
                    $NewServerNameComboBox.enabled = $false
                    $NewPortNameLabel.enabled = $false
                    $NewPortNameBox.enabled = $false
                    $NewPortValueLabel.enabled = $false
                    $NewPortValueBox.enabled = $false
                    $NewDriverNameLabel.enabled = $false
                    $NewDriverNameComboBox.enabled = $false
                    $NewAssetTagLabel.enabled = $false
                    $NewAssetTagBox.enabled = $false
                    $NewSerialNoLabel.enabled = $false
                    $NewSerialNoBox.enabled = $false
                    $NewTapLabel.enabled = $false
                    $NewTapBox.enabled = $false
                    $NewLocationLabel.enabled = $false
                    $NewLocationBox.enabled = $false
                    $NewServerNameListAllCheckBox.enabled = $false
                    $NewServerNameListAllLabel.enabled = $false
                    $NewPrinterProvisionPrinterButton.Enabled = $false
            }
        
    }
            $NewPrinterSearchButton.Enabled = $true
            $newProgressBar.Visible = $false
            $NewIPAddressBox.Enabled = $true
}#CreateNewPrinterSearchClick end

$CreateNewPrinterServerIndexChanged = {

    $ServerName = ($NewServerNameComboBox.SelectedItem).Split(":") | select -Index 0
    $NewDriverList = get-ciminstance -class win32_PrinterDriver -ComputerName $ServerName | select @{n="DriverName";e={(($_.name).split(",") | select -index 0)}} | select -unique -ExpandProperty DriverName
    $NewDriverNameComboBox.items.Clear()
    $pleaseselect = "<Please Select>"    
    $NewDriverNameComboBox.items.add($pleaseselect)  
    $NewDriverNameComboBox.items.addrange(@($NewDriverList))
    $NewDriverNameComboBox.SelectedItem = $pleaseselect
}#CreateNewPrinterServerIndexChanged end

$CreateNewPrinterCopyToClipboardClick = {

[System.Windows.Forms.Clipboard]::SetText($NewPrinterProvisionResultsTextBox.Text)

}#CreateNewPrinterCopyToClipboardClick end

$CreateNewPrinterResetClick = {

    $NewIPAddressLabel.enabled = $true
    $NewIPAddressBox.enabled = $true
    $NewIPAddressBox.Text = $null
    $NewPrinterProvisionResultsTextBox.Enabled = $true
    $NewMacAddressLabel.enabled = $false 
    $NewMacAddressBox.enabled = $false 
    $NewMacAddressBox.Text = $null
    $NewPrinterNameLabel.enabled = $false
    $NewPrinterNameBox.enabled = $false
    $NewPrinterNameBox.Text = $null
    $NewServerNameLabel.enabled = $false
    $NewServerNameComboBox.enabled = $false
    $NewServerNameComboBox.items.Clear()
    $NewPortNameLabel.enabled = $false
    $NewPortNameBox.enabled = $false
    $NewPortNameBox.text = $null
    $NewPortValueLabel.enabled = $false
    $NewPortValueBox.enabled = $false
    $NewPortValueBox.Text = $null
    $NewDriverNameLabel.enabled = $false
    $NewDriverNameComboBox.enabled = $false
    $NewDriverNameComboBox.items.Clear()
    $NewAssetTagLabel.enabled = $false
    $NewAssetTagBox.enabled = $false
    $NewAssetTagBox.Text = $null
    $NewSerialNoLabel.enabled = $false
    $NewSerialNoBox.enabled = $false
    $NewSerialNoBox.text = $null
    $NewTapLabel.enabled = $false
    $NewTapBox.enabled = $false
    $NewTapBox.text = $null
    $NewLocationLabel.enabled = $false
    $NewLocationBox.enabled = $false
    $NewLocationBox.Text = $null
    $NewMakeModelLabel.enabled = $false
    $NewMakeModelBox.Enabled = $false
    $NewMakeModelBox.Text = $null
    $NewServerNameListAllCheckBox.enabled = $false
    $NewServerNameListAllCheckBox.Checked = $false
    $NewServerNameListAllLabel.enabled = $false
    $NewPrinterProvisionResultsTextBox.text = $null
    $NewPrinterProvisionPrinterButton.Enabled = $false
    $NewCopyToClipboardButton.enabled = $false
    $NewResetButton.enabled = $false
    $NewPrinterSearchButton.enabled = $true

}#CreateNewPrinterResetClick end

$CreateNewPrinterProvisionClick = {
if ($NewLocationBox.text -notmatch '[a-zA-Z][a-zA-Z][a-zA-Z]/..') {

    [System.Windows.Forms.MessageBox]::Show("The location field needs to be in the format of SITECODE/BUILDING/FLOOR/ROOM.")


}elseif (($NewPrinterNameBox.Text -ne "") -and ($NewMacAddressBox.text -ne "") -and ($NewServerNameComboBox.SelectedItem -ne "") -and ($NewAssetTagBox.text -ne "") -and ($NewSerialNoBox.text -ne "") -and ($NewTapBox.text -ne "") -and ($NewLocationBox.text -ne "") -and ($NewMakeModelBox.text -ne "") -and ($NewDriverNameComboBox.SelectedItem -notlike "*Please Select*") ) {
        ###DHCP reservation creation
        $NewIPAddressLabel.enabled = $false
        $NewIPAddressBox.enabled = $false
        $NewPrinterProvisionResultsTextBox.Enabled = $true
        $NewMacAddressLabel.enabled = $false 
        $NewMacAddressBox.enabled = $false 
        $NewPrinterNameLabel.enabled = $false
        $NewPrinterNameBox.enabled = $false
        $NewServerNameLabel.enabled = $false
        $NewServerNameComboBox.enabled = $false
        $NewPortNameLabel.enabled = $false
        $NewPortNameBox.enabled = $false
        $NewPortValueLabel.enabled = $false
        $NewPortValueBox.enabled = $false
        $NewDriverNameLabel.enabled = $false
        $NewDriverNameComboBox.enabled = $false
        $NewAssetTagLabel.enabled = $false
        $NewAssetTagBox.enabled = $false
        $NewSerialNoLabel.enabled = $false
        $NewSerialNoBox.enabled = $false
        $NewTapLabel.enabled = $false
        $NewTapBox.enabled = $false
        $NewLocationLabel.enabled = $false
        $NewLocationBox.enabled = $false
        $NewMakeModelLabel.enabled = $false
        $NewMakeModelBox.Enabled = $false
        $NewServerNameListAllCheckBox.enabled = $false
        $NewServerNameListAllLabel.enabled = $false
        $NewPrinterProvisionPrinterButton.Enabled = $false
        $NewPrinterSearchButton.Enabled = $false


        $dhcpservers = @()
        $scope = @()
        $dhcpserverslist = get-content "$PSScriptRoot\DHCPServers.txt"
        
        foreach ($dhcpserverlist in $dhcpserverslist){
    
            $dhcpservers += Get-ciminstance win32_pingstatus -Filter "address='$dhcpserverlist'" | Where-Object {$_.StatusCode -eq 0} |  select -ExpandProperty address
    
        }

       
        [string]$macaddress = ($NewMacAddressBox.text).replace("-","")
        [string]$comment = "Asset:" + $NewAssetTagBox.text + " Serial:" + $NewSerialNoBox.text + " Tap:" + $NewTapBox.text
        $comment
        [string]$ipaddress = $NewIPAddressBox.text
        foreach ($dh in $dhcpservers){
        $scope += Invoke-wmiMethod -Class PS_DhcpServerv4Lease -Namespace root/microsoft/windows/dhcp -ComputerName $dh -Argument $null,$ipaddress -Name GetByIPAddress | select -expand properties | select -expand value | select -first 1 -ExpandProperty ScopeId
        }
        $scope = $scope | select -First 1
    
        [string]$printername = $newPrinterNameBox.text

        foreach ($dh in $dhcpServers){

            $dhcpservernull = Invoke-wmiMethod -Class PS_DhcpServerv4scope -Namespace root/microsoft/windows/dhcp -ComputerName $dh -name get -ArgumentList $null, $scope

            if ($dhcpservernull -ne $null){

                Invoke-wmiMethod -Class PS_DhcpServerv4Reservation -Namespace root/microsoft/windows/dhcp -ComputerName $dh -name add -ArgumentList $macaddress, $dh, $comment, $ipaddress, $printername, $null, $scope, $null

                $dhcpReservationCheck =  Invoke-wmiMethod -Class PS_DhcpServerv4Lease -Namespace root/microsoft/windows/dhcp -ComputerName $dh -Argument $null,$ipaddress -Name GetByIPAddress | select -expand properties | select -expand value | select -first 1 -ExpandProperty addressstate
            
                if ($dhcpReservationCheck -like "*reservation*"){

                    $NewPrinterProvisionResultsTextBox.Text += "Creating DHCP reservation on server " + $dh + " with ip address $ipaddress" + ".......success`r`n"

                } else{$NewPrinterProvisionResultsTextBox.Text += "Creating DHCP reservation on server " + $dh + ".......failed`r`n" }
            }
        }

                #   Printer Port Creation

                $portname = $NewPortNameBox.Text
                $portvalue = $NewPortValueBox.Text
                $ServerName = ($NewServerNameComboBox.SelectedItem).Split(":") | select -Index 0

                $NewPrinterProvisionResultsTextBox.Text += "Creating printer port on server " + $servername + "......."

                $portclass = [wmiclass]"\\$ServerName\root\cimv2:Win32_TcpIpPrinterPort"
                $newPort = $portclass.CreateInstance()
                $newport.Name= "$PortName"
                $newport.SNMPEnabled=$false 
                $newport.Protocol=1 
                $newport.HostAddress= "$PortValue" 
                $newport.Put() > $null

                $printerPortCheck = Get-ciminstance Win32_TCPIPPrinterPort -computername $ServerName -Filter "Name='$PrinterName'"

                if ($printerPortCheck -ne $null){
            
                    $NewPrinterProvisionResultsTextBox.Text += "success`r`n"

                }else{

                    $NewPrinterProvisionResultsTextBox.Text += "failed`r`n"

                }

                #    Printer Queue Creation
                $NewPrinterProvisionResultsTextBox.Text += "Creating printer queue on server " + $servername + "......."

                $DriverName = $NewDriverNameComboBox.SelectedItem
                $PrinterName = $NewPrinterNameBox.Text
                $Comment = "Asset:" + $NewAssetTagBox.text + " Serial: " + $NewSerialNoBox.text + " Tap: " + $NewTapBox.text + " MakeModel: " + $NewMakeModelBox.Text
                $Location = $NewLocationBox.Text 

                $print = "\\$ServerName\root\cimv2:Win32_Printer"
                $print = [WMICLASS]$print
                $print.scope.options.enableprivileges = $true
                $newprinter = $print.createInstance() 
                $newprinter.drivername = "$DriverName"
                $newprinter.PortName = "$PortName"
                $newprinter.Shared = $true
                $newprinter.network = $true
                $newprinter.Sharename = $PrinterName
                $newprinter.Location = "$Location"
                $newprinter.Comment = $Comment
                $newprinter.DeviceID = $PrinterName
                $newprinter.Published = $true
                $newprinter.DoCompleteFirst = $true
                $newprinter.EnableXXXDI = $true
                $newprinter.Put() > $null

                $printerCreationCheck = Get-ciminstance Win32_Printer -computername $ServerName -filter "Name='$printername'"

                if ($printerCreationCheck -ne $null){
            
                    $NewPrinterProvisionResultsTextBox.Text += "success`r`n`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "Printer Installed with the following details:`r`n`r`n"

                    $NewPrinterProvisionResultsTextBox.Text += "PrinterName: " + $NewPrinterNameBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "IPAddress: " + $NewPortValueBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "MacAddress: " + $NewMacAddressBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "DriverName: " + $NewDriverNameComboBox.SelectedItem + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "AssetTag: " +  $NewAssetTagBox.text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "Serial: " + $NewSerialNoBox.text + "`r`n" 
                    $NewPrinterProvisionResultsTextBox.Text += "Tap: " + $NewTapBox.text + "`r`n" 
                    $NewPrinterProvisionResultsTextBox.Text += 'Make/Model: ' + $NewMakeModelBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += 'Location: ' + $NewLocationBox.Text + "`r`n"

                    $alllastprinters = @()
                    $alllastprinters = gc '\\UKXXXXXXSXDDC001\C$\CVS\LastPrinters.txt' | select -Last 30
                    $alllastprinters += $NewPrinterNameBox.Text

                    $alllastprinters | Out-File '\\UKXXXXXXSXDDC001\C$\CVS\LastPrinters.txt'

                }else{

                    $NewPrinterProvisionResultsTextBox.Text += "failed`r`n"

                    $NewPrinterProvisionResultsTextBox.Text += "Printer Installed with errors with the following details:`r`n`r`n"

                    $NewPrinterProvisionResultsTextBox.Text += "PrinterName: " + $NewPrinterNameBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "IPAddress: " + $NewPortValueBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "MacAddress: " + $NewMacAddressBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "DriverName: " + $NewDriverNameComboBox.SelectedItem + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "AssetTag: " +  $NewAssetTagBox.text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += "Serial: " + $NewSerialNoBox.text + "`r`n" 
                    $NewPrinterProvisionResultsTextBox.Text += "Tap: " + $NewTapBox.text + "`r`n" 
                    $NewPrinterProvisionResultsTextBox.Text += 'Make/Model: ' + $NewMakeModelBox.Text + "`r`n"
                    $NewPrinterProvisionResultsTextBox.Text += 'Location: ' + $NewLocationBox.Text + "`r`n"

                }

                $NewCopyToClipboardButton.enabled = $true
                $NewResetButton.enabled = $true
        }else{

            [System.Windows.Forms.MessageBox]::Show("Please Fill out all fields correctly before attempting to Provision.")

        }
}#CreateNewPrinterProvisionClick end

$CreateNewPrinterListAllCheckedClick = {
   
    if ($NewServerNameListAllCheckBox.checked -eq $true){
        
        $installedservers = @()
        $n = $NewServerNameComboBox.items.Count
        For($i = 0; $i -lt $n;$i++){
            $installedservers += $NewServerNameComboBox.items[$i]
        }

        $serverlist = @()
        $allservers = @()
        $fulllistservers = @()
        $NewServerNameComboBox.items.Clear()
        $NewServerNameListAllCheckBox.Enabled = $false
        $NewServerNameComboBox.Enabled = $false
        $root = new-object system.DirectoryServices.DirectoryEntry
        $search = new-object system.DirectoryServices.DirectorySearcher($root)
        $search.PageSize = 1000
        $search.Filter = "(&(objectCategory=computer) (name=*SXXXX*))"
        $search.PropertyNamesOnly = $true
        $search.PropertiesToLoad.Add('Name') | Out-Null
        $search.SearchScope = "subtree"
        $searchresult = $search.FindAll()
    
        foreach($result in $searchresult){
            [System.Windows.Forms.Application]::DoEvents()
	        $serverlist += $result.GetDirectoryEntry().Name
        }
         
        $list = Foreach ($s in $installedServers){ (($s).split(" ") | select -First 1) }           
        $allservers = Compare-Object -referenceObject $serverlist -differenceObject $list | sort inputobject | select @{n="input";e={$_.inputobject + ": 0 printers installed"}} | select -ExpandProperty input

        $allservers += $installedservers
        $NewServerNameComboBox.items.addrange(@($allservers))
        $NewServerNameComboBox.SelectedIndex = 0
        $NewServerNameListAllCheckBox.Enabled = $true
        $NewServerNameComboBox.Enabled = $true
       

    }

    if ($NewServerNameListAllCheckBox.checked -eq $false){ 
        $installedservers = @()
        $n = $NewServerNameComboBox.items.Count

        For($i = 0; $i -lt $n;$i++){

            $serverinstalledno = ($NewServerNameComboBox.items[$i]).Split(" ") | select -Index 1

            if ($serverinstalledno -ne "0"){
                $installedservers += $NewServerNameComboBox.items[$i]
            }
        }

        $NewServerNameComboBox.items.Clear()
        $NewServerNameComboBox.items.addrange(@($installedservers))
        $NewServerNameComboBox.SelectedIndex = 0
        
    }

}#CreateNewPrinterListAllClick end

Function Generate-Form{
    Add-Type -AssemblyName system.windows.forms
    Add-Type -AssemblyName System.Drawing

    #build main form
    $mainform = new-object System.Windows.Forms.Form
    $mainform.Width = 1400
    $mainForm.Height = 570
    $mainform.Text = "Printer ToolKit V1.0 developed by Chris Morland"
    $mainform.StartPosition = "CenterScreen"

    $tabcontrol = new-object System.Windows.Forms.TabControl
    $tabcontrol.DataXXXndings.DefaultDataSourceUpdateMode = 0

    $tabcontrol.Location = new-object System.Drawing.Size(15,10)
    $tabcontrol.Size = new-object System.Drawing.Size(1355,480)
    $tabcontrol.Text = "Search Single Printer"
    $tabcontrol.Name = "tabcontrol"
    $mainform.controls.Add($tabcontrol)


    #############################################################################
    #Tab Create: Search Single Printer
    #############################################################################
    $SearchSinglePrinter =  New-Object System.Windows.Forms.TabPage
    $SearchSinglePrinter.UseVisualStyleBackColor = $true
    $SearchSinglePrinter.Size = new-object System.Drawing.Size(1355,500)
    $SearchSinglePrinter.Text = "Search Single Printer"
    $SearchSinglePrinter.Name = "SearchSinglePrinter"
    $tabcontrol.Controls.Add($SearchSinglePrinter)

    #Add controls Search Single Printer
    
    $SearchSinglePrinterNameLabel = new-object System.Windows.Forms.Label
    $SearchSinglePrinterNameLabel.Location = new-object System.Drawing.Size(560,27)
    $SearchSinglePrinterNameLabel.Text = "PrinterName:"
    
    $SearchSinglePrinterNameBox = new-object System.Windows.Forms.TextBox
    $SearchSinglePrinterNameBox.Location = new-object System.Drawing.Size(630,23)
    $SearchSinglePrinterNameBox.Size = new-object System.Drawing.Size(150,23)
 
    $SearchSingleFindButton = new-object System.Windows.Forms.Button
    $SearchSingleFindButton.Location = new-object System.Drawing.Size(810,23)
    $SearchSingleFindButton.Size = new-object System.Drawing.Size(40,19)
    $SearchSingleFindButton.Text = "Find"

    $SearchSinglegridview = new-object System.Windows.Forms.DataGridView
    $SearchSinglegridview.Location = new-object System.Drawing.Size(15,70) 
    $SearchSinglegridview.Size = new-object System.Drawing.Size(1317,100)
    $SearchSinglegridview.ScrollBars
    $SearchSinglegridview.ColumnCount = 10
    $SearchSinglegridview.ColumnHeadersVisible = $true
    $SearchSinglegridview.Columns[0].Name = "ServerName"
    $SearchSinglegridview.Columns[1].Name = "PrinterName"
    $SearchSinglegridview.Columns[2].Name = "Status"
    $SearchSinglegridview.Columns[3].Name = "PingCheck"
    $SearchSinglegridview.Columns[4].Name = "JobsInQueue"
    $SearchSinglegridview.Columns[5].Name = 'PortName'
    $SearchSinglegridview.Columns[6].Name = 'PortValue'
    $SearchSinglegridview.Columns[7].Name = 'DriverName'
    $SearchSinglegridview.Columns[8].Name = "Comment"
    $SearchSinglegridview.Columns[9].Name = "Location"
    
    
    $SearchSinglegridview.autoresizecolumns()
    $SearchSinglegridview.AutoResizeRows()
    $SearchSinglegridview.SelectionMode = 'FullRowSelect'
    $SearchSinglegridview.MultiSelect = $false
    $SearchSinglegridview.ReadOnly = $true
    $SearchSinglegridview.DataSource = $null
    
    $SearchSingleClearQueueButton = new-object System.Windows.Forms.Button
    $SearchSingleClearQueueButton.Location = new-object System.Drawing.Size(15,200)
    $SearchSingleClearQueueButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleClearQueueButton.Text = "Clear Print Queue"
    $SearchSingleClearQueueButton.Enabled = $false

    $SearchSingleSendTestPrintButton = new-object System.Windows.Forms.Button
    $SearchSingleSendTestPrintButton.Location = new-object System.Drawing.Size(140,200)
    $SearchSingleSendTestPrintButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleSendTestPrintButton.Text = "Send Test Print"
    $SearchSingleSendTestPrintButton.Enabled = $false

    $SearchSingleEditPrinterQueueButton = new-object System.Windows.Forms.Button
    $SearchSingleEditPrinterQueueButton.Location = new-object System.Drawing.Size(265,200)
    $SearchSingleEditPrinterQueueButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleEditPrinterQueueButton.Text = "Edit Printer Queue"
    $SearchSingleEditPrinterQueueButton.Enabled = $false

    $SearchSingleRecreatePrinterQueueButton = new-object System.Windows.Forms.Button
    $SearchSingleRecreatePrinterQueueButton.Location = new-object System.Drawing.Size(390,200)
    $SearchSingleRecreatePrinterQueueButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleRecreatePrinterQueueButton.Text = "Recreate Queue"
    $SearchSingleRecreatePrinterQueueButton.Enabled = $false

    $SearchSingleCheckDHCPButton = new-object System.Windows.Forms.Button
    $SearchSingleCheckDHCPButton.Location = new-object System.Drawing.Size(515,200)
    $SearchSingleCheckDHCPButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleCheckDHCPButton.Text = "Check DHCP"
    $SearchSingleCheckDHCPButton.Enabled = $false
    
    $SearchSingleRDPToXXXXButton = new-object System.Windows.Forms.Button
    $SearchSingleRDPToXXXXButton.Location = new-object System.Drawing.Size(640,200)
    $SearchSingleRDPToXXXXButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleRDPToXXXXButton.Text = "RDP to Server"
    $SearchSingleRDPToXXXXButton.Enabled = $false


    $SearchSingleDeletePrinterQueueButton = new-object System.Windows.Forms.Button
    $SearchSingleDeletePrinterQueueButton.Location = new-object System.Drawing.Size(1090,200)
    $SearchSingleDeletePrinterQueueButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleDeletePrinterQueueButton.Text = "Delete Printer Queue"
    $SearchSingleDeletePrinterQueueButton.Enabled = $false

    $SearchSingleRestartSpoolerButton = new-object System.Windows.Forms.Button
    $SearchSingleRestartSpoolerButton.Location = new-object System.Drawing.Size(1215,200)
    $SearchSingleRestartSpoolerButton.Size = new-object System.Drawing.Size(120,60)
    $SearchSingleRestartSpoolerButton.Text = "Restart Server Spooler Service"
    $SearchSingleRestartSpoolerButton.Enabled = $false
    
    
    $SearchSinglePrinter.controls.Add($SearchSinglePrinterNameBox)
    $SearchSinglePrinter.controls.Add($SearchSinglePrinterNameLabel)
    $SearchSinglePrinter.controls.Add($SearchSingleFindButton)
    $SearchSinglePrinter.controls.Add($SearchSingleGridView)
    $SearchSinglePrinter.controls.Add($SearchSingleClearQueueButton)
    $SearchSinglePrinter.controls.Add($SearchSingleSendTestPrintButton)
    $SearchSinglePrinter.controls.Add($SearchSingleEditPrinterQueueButton)
    $SearchSinglePrinter.controls.Add($SearchSingleRecreatePrinterQueueButton)
    $SearchSinglePrinter.controls.Add($SearchSingleDeletePrinterQueueButton)
    $SearchSinglePrinter.controls.Add($SearchSingleRestartSpoolerButton)
    $SearchSinglePrinter.controls.Add($SearchSingleCheckDHCPButton)
    $SearchSinglePrinter.controls.Add($SearchSingleRDPToXXXXButton)

    #Add Events Search Single Printer
    $SearchSingleFindButton.Add_Click( $SearchSingleFindButtonClick )
    $SearchSingleClearQueueButton.Add_Click( $SearchSingleClearQueueClick )
    $SearchSingleSendTestPrintButton.Add_Click( $SearchSingleSendTestPrintClick )
    $SearchSingleEditPrinterQueueButton.Add_Click( $SearchSingleEditPrinterQueueClick )
    $SearchSingleRDPToXXXXButton.Add_Click( $SearchSingleRDPToXXXXClick )

    #############################################################################
    #Tab Create: Create New Printer
    #############################################################################    
    $CreateNewPrinter =  New-Object System.Windows.Forms.TabPage
    $CreateNewPrinter.UseVisualStyleBackColor = $true
    $CreateNewPrinter.Size = new-object System.Drawing.Size(1455,600)
    $CreateNewPrinter.Text = "Create New Printer"
    $CreateNewPrinter.Name = "CreateNewPrinter"
    $tabcontrol.Controls.Add($CreateNewPrinter)


   
    #Add Controls

    $NewIPAddressLabel = new-object System.Windows.Forms.Label
    $NewIPAddressLabel.Location = new-object System.Drawing.Size(15,27)
    $NewIPAddressLabel.Size = new-object System.Drawing.Size(73,23)
    $NewIPAddressLabel.Text = "IP Address:"

    $NewIPAddressBox = new-object System.Windows.Forms.TextBox
    $NewIPAddressBox.Location = new-object System.Drawing.Size(90,23)
    $NewIPAddressBox.Size = new-object System.Drawing.Size(100,23)
    $NewIPAddressBox.Visible = $true

    $NewProgressBar = new-object System.Windows.Forms.ProgressBar
    $NewProgressBar.Name = "ProgressBar1"
    $NewProgressBar.Value = 0
    $NewProgressBar.Style = "Continuous"
    $NewProgressBar.Location = new-object System.Drawing.Size(400,420)
    $NewProgressBar.Size = new-object System.Drawing.Size(500,15)
    $NewProgressBar.Visible = $false

    $NewMacAddressLabel = new-object System.Windows.Forms.Label
    $NewMacAddressLabel.Location = new-object System.Drawing.Size(400,27)
    $NewMacAddressLabel.Size = new-object System.Drawing.Size(73,23)
    $NewMacAddressLabel.Text = "MacAddress:"
    $NewMacAddressLabel.Visible = $true
    $NewMacAddressLabel.Enabled = $false
    
    $NewMacAddressBox = new-object System.Windows.Forms.TextBox
    $NewMacAddressBox.Location = new-object System.Drawing.Size(480,23)
    $NewMacAddressBox.Size = new-object System.Drawing.Size(200,23)
    $NewMacAddressBox.Visible = $true
    $NewMacAddressBox.ReadOnly = $true
    $NewMacAddressBox.Enabled = $false

    $NewPrinterNameLabel = new-object System.Windows.Forms.Label
    $NewPrinterNameLabel.Location = new-object System.Drawing.Size(400,57)
    $NewPrinterNameLabel.Size = new-object System.Drawing.Size(73,23)
    $NewPrinterNameLabel.Text = "PrinterName:"
    $NewPrinterNameLabel.Visible = $true
    $NewPrinterNameLabel.Enabled = $false
    
    $NewPrinterNameBox = new-object System.Windows.Forms.TextBox
    $NewPrinterNameBox.Location = new-object System.Drawing.Size(480,53)
    $NewPrinterNameBox.Size = new-object System.Drawing.Size(200,23)
    $NewPrinterNameBox.Visible = $true
    $NewPrinterNameBox.ReadOnly = $true
    $NewPrinterNameBox.Enabled = $false

    $NewServerNameLabel = new-object System.Windows.Forms.Label
    $NewServerNameLabel.Location = new-object System.Drawing.Size(400,87)
    $NewServerNameLabel.Size = new-object System.Drawing.Size(73,23)
    $NewServerNameLabel.Text = "ServerName:"
    $NewServerNameLabel.Visible = $true
    $NewServerNameLabel.Enabled = $false

    $NewServerNameComboBox = new-object System.Windows.Forms.ComboBox
    $NewServerNameComboBox.Location = new-object System.Drawing.Size(480,83)
    $NewServerNameComboBox.Size = new-object System.Drawing.Size(200,23)
    $NewServerNameComboBox.XXXndingContext = $CreateNewPrinter.XXXndingContext
    $NewServerNameComboBox.DataSource = $ServerListFinal
    $NewServerNameComboBox.Visible = $true
    $NewServerNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $NewServerNameComboBox.Enabled = $false
    
    $NewServerNameListAllCheckBox = new-object System.Windows.Forms.CheckBox
    $NewServerNameListAllCheckBox.Location = new-object System.Drawing.Size(686,83)
    $NewServerNameListAllCheckBox.Size = new-object System.Drawing.Size(13,20)
    $NewServerNameListAllCheckBox.Visible = $true
    $NewServerNameListAllCheckBox.Enabled = $false

    $NewServerNameListAllLabel = new-object System.Windows.Forms.Label
    $NewServerNameListAllLabel.Location = new-object System.Drawing.Size(700,87)
    $NewServerNameListAllLabel.Size = new-object System.Drawing.Size(73,23)
    $NewServerNameListAllLabel.Text = "ListAll"
    $NewServerNameListAllLabel.Visible = $true
    $NewServerNameListAllLabel.Enabled = $false

    $NewPortNameLabel = new-object System.Windows.Forms.Label
    $NewPortNameLabel.Location = new-object System.Drawing.Size(400,117)
    $NewPortNameLabel.Size = new-object System.Drawing.Size(73,23)
    $NewPortNameLabel.Text = "PortName:"
    $NewPortNameLabel.Visible = $true
    $NewPortNameLabel.Enabled = $false

    $NewPortNameBox = new-object System.Windows.Forms.TextBox
    $NewPortNameBox.Location = new-object System.Drawing.Size(480,113)
    $NewPortNameBox.Size = new-object System.Drawing.Size(200,23)
    $NewPortNameBox.ReadOnly = $true
    $NewPortNameBox.Visible = $true
    $NewPortNameBox.Enabled = $false

    $NewPortValueLabel = new-object System.Windows.Forms.Label
    $NewPortValueLabel.Location = new-object System.Drawing.Size(400,147)
    $NewPortValueLabel.Size = new-object System.Drawing.Size(73,23)
    $NewPortValueLabel.Text = "PortValue:"
    $NewPortValueLabel.Visible = $true
    $NewPortValueLabel.Enabled = $false

    $NewPortValueBox = new-object System.Windows.Forms.TextBox
    $NewPortValueBox.Location = new-object System.Drawing.Size(480,143)
    $NewPortValueBox.Size = new-object System.Drawing.Size(200,23)
    $NewPortValueBox.Text = $PortValue
    $NewPortValueBox.ReadOnly = $true
    $NewPortValueBox.Visible = $true
    $NewPortValueBox.Enabled = $false

    $NewDriverNameLabel = new-object System.Windows.Forms.Label
    $NewDriverNameLabel.Location = new-object System.Drawing.Size(400,177)
    $NewDriverNameLabel.Size = new-object System.Drawing.Size(73,23)
    $NewDriverNameLabel.Text = "DriverName:"
    $NewDriverNameLabel.Visible = $true
    $NewDriverNameLabel.Enabled = $false

    $NewDriverNameComboBox = new-object System.Windows.Forms.ComboBox
    $NewDriverNameComboBox.Location = new-object System.Drawing.Size(480,173)
    $NewDriverNameComboBox.Size = new-object System.Drawing.Size(200,23)
    $NewDriverNameComboBox.XXXndingContext = $CreateNewPrinter.XXXndingContext
    $NewDriverNameComboBox.DataSource = $NewDriverList
    $NewDriverNameComboBox.Visible = $true
    $NewDriverNameComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $NewDriverNameComboBox.Enabled = $false

    $NewAssetTagLabel = new-object System.Windows.Forms.Label
    $NewAssetTagLabel.Location = new-object System.Drawing.Size(400,207)
    $NewAssetTagLabel.Size = new-object System.Drawing.Size(73,23)
    $NewAssetTagLabel.Text = "AssetTag:"
    $NewAssetTagLabel.Visible = $true
    $NewAssetTagLabel.Enabled = $false

    $NewAssetTagBox = new-object System.Windows.Forms.TextBox
    $NewAssetTagBox.Location = new-object System.Drawing.Size(480,203)
    $NewAssetTagBox.Size = new-object System.Drawing.Size(200,23)
    $NewAssetTagBox.Visible = $true
    $NewAssetTagBox.Enabled = $false

    $NewSerialNoLabel = new-object System.Windows.Forms.Label
    $NewSerialNoLabel.Location = new-object System.Drawing.Size(400,237)
    $NewSerialNoLabel.Size = new-object System.Drawing.Size(73,23)
    $NewSerialNoLabel.Text = "SerialNo:"
    $NewSerialNoLabel.Visible = $true
    $NewSerialNoLabel.Enabled = $false

    $NewSerialNoBox = new-object System.Windows.Forms.TextBox
    $NewSerialNoBox.Location = new-object System.Drawing.Size(480,233)
    $NewSerialNoBox.Size = new-object System.Drawing.Size(200,23)
    $NewSerialNoBox.Visible = $true
    $NewSerialNoBox.Enabled = $false

    $NewTapLabel = new-object System.Windows.Forms.Label
    $NewTapLabel.Location = new-object System.Drawing.Size(400,267)
    $NewTapLabel.Size = new-object System.Drawing.Size(73,23)
    $NewTapLabel.Text = "Tap:"
    $NewTapLabel.Visible = $true
    $NewTapLabel.Enabled = $false

    $NewTapBox = new-object System.Windows.Forms.TextBox
    $NewTapBox.Location = new-object System.Drawing.Size(480,263)
    $NewTapBox.Size = new-object System.Drawing.Size(200,23)
    $NewTapBox.Visible = $true
    $NewTapBox.Enabled = $false

    $NewLocationLabel = new-object System.Windows.Forms.Label
    $NewLocationLabel.Location = new-object System.Drawing.Size(400,297)
    $NewLocationLabel.Size = new-object System.Drawing.Size(73,23)
    $NewLocationLabel.Text = "Location:"
    $NewLocationLabel.Visible = $true
    $NewLocationLabel.Enabled = $false

    $NewLocationBox = new-object System.Windows.Forms.TextBox
    $NewLocationBox.Location = new-object System.Drawing.Size(480,293)
    $NewLocationBox.Size = new-object System.Drawing.Size(200,23)
    $NewLocationBox.Visible = $true
    $NewLocationBox.Enabled = $false

    $NewMakeModelLabel = new-object System.Windows.Forms.Label
    $NewMakeModelLabel.Location = new-object System.Drawing.Size(400,327)
    $NewMakeModelLabel.Size = new-object System.Drawing.Size(73,23)
    $NewMakeModelLabel.Text = "Make/Model:"
    $NewMakeModelLabel.Visible = $true
    $NewMakeModelLabel.Enabled = $false

    $NewMakeModelBox = new-object System.Windows.Forms.TextBox
    $NewMakeModelBox.Location = new-object System.Drawing.Size(480,323)
    $NewMakeModelBox.Size = new-object System.Drawing.Size(200,23)
    $NewMakeModelBox.Visible = $true
    $NewMakeModelBox.Enabled = $false

    $NewPrinterSearchButton = new-object System.Windows.Forms.Button
    $NewPrinterSearchButton.Location = new-object System.Drawing.Size(200,23)
    $NewPrinterSearchButton.Size = new-object System.Drawing.Size(70,20)
    $NewPrinterSearchButton.Text = "Search"

    $NewPrinterProvisionPrinterButton = new-object System.Windows.Forms.Button
    $NewPrinterProvisionPrinterButton.Location = new-object System.Drawing.Size(520,353)
    $NewPrinterProvisionPrinterButton.Size = new-object System.Drawing.Size(100,35)
    $NewPrinterProvisionPrinterButton.Text = "Provision New Printer"
    $NewPrinterProvisionPrinterButton.Visible = $true
    $NewPrinterProvisionPrinterButton.Enabled = $false

    $NewPrinterProvisionResultsTextBox = new-object System.Windows.Forms.TextBox
    $NewPrinterProvisionResultsTextBox.Location = new-object System.Drawing.Size(800,23)
    $NewPrinterProvisionResultsTextBox.Size = new-object System.Drawing.Size(500,323)
    $NewPrinterProvisionResultsTextBox.Multiline = $true
    $NewPrinterProvisionResultsTextBox.Visible = $true
    $NewPrinterProvisionResultsTextBox.Enabled = $false
    $NewPrinterProvisionResultsTextBox.ReadOnly = $true
    $NewPrinterProvisionResultsTextBox.SelectionLength = 0

    $NewCopyToClipboardButton = new-object System.Windows.Forms.Button
    $NewCopyToClipboardButton.Location = new-object System.Drawing.Size(1100,353)
    $NewCopyToClipboardButton.Size = new-object System.Drawing.Size(100,35)
    $NewCopyToClipboardButton.Text = "Copy To Clipboard"
    $NewCopyToClipboardButton.Visible = $true
    $NewCopyToClipboardButton.Enabled = $false

    $NewResetButton = new-object System.Windows.Forms.Button
    $NewResetButton.Location = new-object System.Drawing.Size(1200,353)
    $NewResetButton.Size = new-object System.Drawing.Size(100,35)
    $NewResetButton.Text = "Clear Screen to Do Another"
    $NewResetButton.Visible = $true
    $NewResetButton.Enabled = $false

    $CreateNewPrinter.controls.Add($NewIPAddressLabel)
    $CreateNewPrinter.controls.Add($NewIPAddressBox)
    $CreateNewPrinter.controls.Add($NewMacAddressLabel)
    $CreateNewPrinter.controls.Add($NewMacAddressBox)
    $CreateNewPrinter.controls.Add($NewPrinterNameLabel)
    $CreateNewPrinter.controls.Add($NewPrinterNameBox)
    $CreateNewPrinter.controls.Add($NewServerNameLabel)
    $CreateNewPrinter.controls.Add($NewServerNameComboBox)
    $CreateNewPrinter.controls.Add($NewPortNameLabel)
    $CreateNewPrinter.controls.Add($NewPortNameBox)
    $CreateNewPrinter.controls.Add($NewPortValueLabel)
    $CreateNewPrinter.controls.Add($NewPortValueBox)
    $CreateNewPrinter.controls.Add($NewDriverNameLabel)
    $CreateNewPrinter.controls.Add($NewDriverNameComboBox)
    $CreateNewPrinter.controls.Add($NewAssetTagLabel)
    $CreateNewPrinter.controls.Add($NewAssetTagBox)
    $CreateNewPrinter.controls.Add($NewSerialNoLabel)
    $CreateNewPrinter.controls.Add($NewSerialNoBox)
    $CreateNewPrinter.controls.Add($NewTapLabel)
    $CreateNewPrinter.controls.Add($NewTapBox)
    $CreateNewPrinter.controls.Add($NewLocationLabel)
    $CreateNewPrinter.controls.Add($NewLocationBox)
    $CreateNewPrinter.controls.Add($NewMakeModelLabel)
    $CreateNewPrinter.controls.Add($NewMakeModelBox)
    $CreateNewPrinter.controls.Add($NewPrinterSearchButton)
    $CreateNewPrinter.controls.Add($NewProgressBar)
    $CreateNewPrinter.controls.Add($NewPrinterProvisionPrinterButton)
    $CreateNewPrinter.controls.Add($NewServerNameListAllCheckBox)
    $CreateNewPrinter.controls.Add($NewServerNameListAllLabel)
    $CreateNewPrinter.controls.Add($NewPrinterProvisionResultsTextBox)
    $CreateNewPrinter.controls.Add($NewCopyToClipboardButton)
    $CreateNewPrinter.controls.Add($NewResetButton)

    #Add Events
    $NewPrinterSearchButton.Add_Click( $CreateNewPrinterSearchClick )
    $NewPrinterProvisionPrinterButton.Add_Click( $CreateNewPrinterProvisionClick )
    $NewServerNameComboBox.Add_SelectedIndexChanged( $CreateNewPrinterServerIndexChanged )
    $NewCopyToClipboardButton.Add_Click( $CreateNewPrinterCopyToClipboardClick )
    $NewResetButton.Add_Click( $CreateNewPrinterResetClick )
    $NewServerNameListAllCheckBox.Add_Click( $CreateNewPrinterListAllCheckedClick )
    
    $mainForm.ShowDialog()
    $mainForm.Close()
    $mainform.Dispose()
}
Generate-Form | Out-Null