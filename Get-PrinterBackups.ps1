$stringdate = (Get-Date -Format dd-MM-yy).ToString()
$global:dhcpserverslist = ""
$global:PrintServerWildcard = ""

Function Test-Connection2 {
	Param(

        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline = $true)]
	    $ToIPAddressorDNSName,
	
        [Parameter(Mandatory=$false,Position=1)]
	    $FromIPAddressorDNSName = ‘LocalHost’

	)

    $result = [WMI]('\\' + $FromIPAddressorDNSName + '\root\cimv2:win32_pingstatus.Address="' + $ToIPAddressorDNSName + '"')

    If ($result.statuscode –eq 0){

    $status = "Responding"

    }else{$status = "TimedOut"}

        if($result.Address -ne $result.ProtocolAddress){
           $ipaddress = $result.ProtocolAddress;
        }else {$ipaddress = "n/a"}
    
    $object = [PSCustomObject]@{

        'Source' = $result.__SERVER;
        'DestinationName' = $result.Address;
        'DestinationIP' = $ipaddress;
        'Status' = "$status";
        'ResponseTime' = $result.ResponseTime;
        'BufferSize' = $result.BufferSize;
        'ReplySize' = $result.ReplySize;
        'ReplyInconsistency' = $result.ReplyInconsistency;
        'ResponseTimeToLive' = $result.ResponseTimeToLive;
        'TimeToLive' = $result.TimeToLive

    }

    $object.PSObject.TypeNames.Insert(0, 'User.Information')
    $defaultdisplayset = 'Source','DestinationName','DestinationIP','Status'
    $defaultdisplaypropertyset = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultdisplayset)
    $psstandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultdisplaypropertyset)
    $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    $object

}
        
Function Get-ADServerList{
        
        Param(

        [Parameter(Mandatory=$true,Position=0)]
	    $ServerName
        )
        
        $root = new-object system.DirectoryServices.DirectoryEntry
        $search = new-object system.DirectoryServices.DirectorySearcher($root)
        $search.PageSize = 1000
        $search.Filter = "(&(objectCategory=computer) (name=$ServerName))"
        $search.PropertyNamesOnly = $true
        $search.PropertiesToLoad.Add('Name') | Out-Null
        $search.SearchScope = "subtree"
        $searchresult = $search.FindAll()
    
        foreach($result in $searchresult){
            [System.Windows.Forms.Application]::DoEvents()
	        $serverlist += $result.GetDirectoryEntry().Name
        }
        $serverlist
}

$codecontainer = {
    Param(
        [parameter( Mandatory=$true, Position=0, ValueFromPipeline = $true)]
        [string] $computerName
    )

    Function Test-Connection2 {
	    Param(

            [Parameter(Mandatory=$true,Position=0,ValueFromPipeline = $true)]
	        $ToIPAddressorDNSName,
	
            [Parameter(Mandatory=$false,Position=1)]
	        $FromIPAddressorDNSName = ‘LocalHost’

	    )

        $result = [WMI]('\\' + $FromIPAddressorDNSName + '\root\cimv2:win32_pingstatus.Address="' + $ToIPAddressorDNSName + '"')

        If ($result.statuscode –eq 0){

        $status = "Responding"

        }else{$status = "TimedOut"}

            if($result.Address -ne $result.ProtocolAddress){
               $ipaddress = $result.ProtocolAddress;
            }else {$ipaddress = "n/a"}
    
        $object = [PSCustomObject]@{

            'Source' = $result.__SERVER;
            'DestinationName' = $result.Address;
            'DestinationIP' = $ipaddress;
            'Status' = "$status";
            'ResponseTime' = $result.ResponseTime;
            'BufferSize' = $result.BufferSize;
            'ReplySize' = $result.ReplySize;
            'ReplyInconsistency' = $result.ReplyInconsistency;
            'ResponseTimeToLive' = $result.ResponseTimeToLive;
            'TimeToLive' = $result.TimeToLive

        }

        $object.PSObject.TypeNames.Insert(0, 'User.Information')
        $defaultdisplayset = 'Source','DestinationName','DestinationIP','Status'
        $defaultdisplaypropertyset = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultdisplayset)
        $psstandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultdisplaypropertyset)
        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
        $object

    }

    Function Get-ADPrinterList{
        
        Param(

            [Parameter(Mandatory=$true,Position=0)]
	        $ServerName
        )

        $PrinterList = @()

        #search active directory for queue server names
        $serverlist = @()
        $root = new-object system.DirectoryServices.DirectoryEntry
        $search = new-object system.DirectoryServices.DirectorySearcher($root)
        $search.PageSize = 1000
        $search.Filter = "(&(objectCategory=printQueue) (ShortServerName=$ServerName))"
        $search.PropertyNamesOnly = $true
        $search.PropertiesToLoad.Add('printername') | Out-Null
        $search.PropertiesToLoad.Add('ShortServerName') | Out-Null
        $search.PropertiesToLoad.Add('Description') | Out-Null
        $search.PropertiesToLoad.Add('PortName') | Out-Null
        $search.PropertiesToLoad.Add('DriverName') | Out-Null
        $search.PropertiesToLoad.Add('Location') | Out-Null
        $search.PropertiesToLoad.Add('WhenCreated') | Out-Null
        $search.PropertiesToLoad.Add('WhenChanged') | Out-Null


        $search.SearchScope = "subtree"
        $searchresult = $search.FindAll()

        foreach($result in $searchresult){
  
            $object = [PSCustomObject]@{

	        'PrinterName' = ($result.GetDirectoryEntry().printername | select -ExpandProperty $_ );
            'ServerName' = ($result.GetDirectoryEntry().ShortServerName | select -ExpandProperty $_ );
            'Description' = ($result.GetDirectoryEntry().Description | select -ExpandProperty $_ );
            'DriverName' = ($result.GetDirectoryEntry().DriverName | select -ExpandProperty $_ );
            'Location' = ($result.GetDirectoryEntry().Location | select -ExpandProperty $_ );
            'PortName' = ($result.GetDirectoryEntry().PortName | select -ExpandProperty $_ );
            'WhenCreated' = ($result.GetDirectoryEntry().WhenCreated | select -ExpandProperty $_ );
            'WhenChanged' = ($result.GetDirectoryEntry().WhenChanged | select -ExpandProperty $_ )

            }
            $object

    }



    }

    Function Get-PrinterPortIPAddress {

        Param(

            [Parameter(Mandatory=$true,Position=0)]
            $PortName,

            [Parameter(Mandatory=$true,Position=1)]
            $ServerName
        )

        [WMI]('\\' + $servername + '\root\cimv2:win32_TCPIPPrinterPort.Name="' + $portname + '"') | select -ExpandProperty HostAddress 

    }

    Function Get-PrinterDHCPAddress {
            Param(
        
            [Parameter(Mandatory = $true, Position=0, ValueFromPipeline = $true)]
            $IPAddress
        
            )
            $dhcpaddress = @()
                
            foreach ($dhcpserverl in $dhcpserverslist){
                
                    $dhcpaddress += Invoke-wmiMethod -Class PS_DhcpServerv4Lease -Namespace root/microsoft/windows/dhcp -ComputerName $dhcpserverl -Argument $null,$IPAddress -Name GetByIPAddress | select -expand properties | select -expand value | select -first 1 HostName,ClientId,AddressState

            }

            return $dhcpaddress

    }

    $printers = Get-ADPrinterList -ServerName $ComputerName | ForEach-Object {

        $portvalue = Get-PrinterPortIPAddress -PortName $_.PortName -ServerName $_.ServerName
        $dhcpaddress = $portvalue | Get-PrinterDHCPAddress
        $reservations = ($dhcpaddress | Where-Object {$_.AddressState -like "*reservation*" } | select -ExpandProperty AddressState ).count
        $hostname1 = $dhcpaddress.HostName | select -Index 0
        $macaddress1 = $dhcpaddress.ClientId | select -index 0
        if( $reservations -ge 2){ 
            $Hostname2 = $dhcpaddress.HostName | select -Index 1
            $macAddress2 = $dhcpaddress.ClientId | select -Index 1
        }else{
            $macaddress2 = $null
            $Hostname2 = $null      
        }
        if ($portvalue -ne $null){
        
            $networkcheck = $portvalue | Test-Connection2 -FromIPAddressorDNSName $ComputerName | select -ExpandProperty status

        }else{$networkcheck = "TimedOut"}
        $macmatch = $macaddress1 -eq $macAddress2

        $issuesfound = @()
        if ($_.PrinterName -notmatch 'XXXXX') { $issuesfound += "PrinterName not correct format"}
        if ( ($_.Description  -notlike "*asset*") -and ($_.Description  -notlike "*serial*") -and ($_.Description  -notlike "*tap*") ) { $issuesfound += "Description not correct format"}
        if ($_.Location -notmatch '[a-zA-Z][a-zA-Z][a-zA-Z]/..') { $issuesfound += "Location not correct format"}
        if ($_.PortName -notmatch 'XXX[a-zA-Z][a-zA-Z][a-zA-Z]\d{4}XXX') { $issuesfound += "PortName not in the correct format"}
        if ($portvalue -notmatch ("^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$") ) { $issuesfound += "PortValue IP address not correct format"}
        if ($reservations -lt 2) { $issuesfound += "DHCP reservations less than 2"}
        if (($hostname1 -notmatch 'XXXX[a-zA-Z][a-zA-Z][a-zA-Z]\d{4}XXX') -or ($hostname2 -notmatch 'XXXX[a-zA-Z][a-zA-Z][a-zA-Z]\d{4}XX') ) { $issuesfound += "NetBios not Printername"}
        if (($hostname1 -notmatch 'XXXXXX') -or ($hostname2 -notmatch 'XXXXXX') ) { $issuesfound += "DHCP not showing FQDN"}

        if ($issuesfound -ne $null){
            
            $migrationready = "No"

        }else{ $migrationready = "Yes"}

        $issuesfound = $issuesfound -join ","

        [pscustomobject]@{
            PrinterName = $_.PrinterName
            MigrationReady = $migrationready
            Issues = (@($issuesfound) | out-string).Trim()
            ServerName = $_.ServerName
            Description = $_.Description
            DriverName = $_.DriverName
            Location = $_.Location
            PortName = $_.PortName
            PortValue = $portvalue
            NetworkCheck = $networkcheck
            NumberOfDHCPReservations = $reservations
            MacAddressesMatch = $macmatch
            DHCPHostName1 = $dhcpaddress.HostName | select -Index 0
            DHCPMacAddress1 = $dhcpaddress.ClientId | select -index 0
            DHCPHostName2 = $Hostname2
            DHCPMacAddress2 = $macaddress2
            WhenCreated = $_.WhenCreated
            WhenChanged = $_.WhenChanged
        }
    }
    return $printers
}
$datecheck = (get-date).adddays(-200)
get-childitem C:\PrinterBackups | Where-Object { $_.creationtime -lt $date } | Remove-Item

$RespondingServers = @()
$TimedOutServers = @()

$ServerList = Get-ADServerList -ServerName *$global:PrintServerWildcard* | select -First 3

foreach ($server in $serverlist) {

    $Ping = Test-Connection2 -ToIPAddressorDNSName $Server

    if ($Ping.Status -eq "Responding"){

        $RespondingServers += $Server

    }else {

        $TimedOutServers += $Server

    }


}

$runspacepool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1,10)
$runspacepool.ApartmentState = "MTA"
$runspacepool.Open()

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$threads = @()
foreach ($c in $RespondingServers) {

    $runspaceobject = [pscustomobject]@{
        
        Runspace = [powershell]::Create()
        Invoker = $null

    }
    $runspaceobject.Runspace.RunspacePool = $runspacepool
    $runspaceobject.Runspace.AddScript($codecontainer) | Out-Null
    $runspaceobject.Runspace.AddArgument($c) | Out-Null
    $runspaceobject.Invoker = $runspaceObject.RunSpace.BeginInvoke()
    $threads += $runspaceobject
    $elapsed = $stopwatch.Elapsed
    Write-Host "Finished creating runspace for $c. Elapsed time : $elapsed"

}

while ($threads.Invoker.IsCompleted -contains $false){}

$elapsed = $stopwatch.Elapsed
Write-Host "All runspaces completed. Elapsed time: $elapsed"

$threadresults = @()
foreach ($t in $threads){

    $threadresults += $t.Runspace.EndInvoke($t.Invoker)
    $t.Runspace.Dispose()
}

$runspacepool.Close()
$runspacepool.Dispose()

$threadresults | Export-Csv "C:\PrinterBackups\AllPrinters_$stringdate.csv" -NoTypeInformation