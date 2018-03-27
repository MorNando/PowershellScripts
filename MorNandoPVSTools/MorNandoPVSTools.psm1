Function Get-mPVSdevice {
<#
.SYNOPSIS
List all citrix PVS devices and their attributes with no parameter. It also has search options using the named parameters.                               
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly. We need object data!                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.

.PARAMETER DeviceName
Provide the name of the device you which to search on. This parameter can only be used by itself but it can accept arrays.

.PARAMETER DeviceMac
Provide the mac address you which to search on. This parameter can only be used by itself but it can accept arrays.

.PARAMETER CollectionName
CollectionName must be used with the sitename parameter. It provides a list of devices and their attributes inside a collection.

.PARAMETER ImageName
ImageName also requires the SiteName and StoreName parameter to work. It provides a list of devices and their attributes which use a specific image.

.PARAMETER SiteName
SiteName is a requirement for both CollectionName and ImageName.

.PARAMETER StoreName
StoreName is a requirement for the ImageName parameter.

.PARAMETER PVSserver
Search for devices by the PVSserver it is using. It doesn't work with other parameters
	
.EXAMPLE	
Get-mPVSdevice
	
This will list all devices with their attributes

.EXAMPLE	
Get-mPVSdevice -DeviceName device1,device2,device3
	
List devices by a name search with their attributes by seperating with a comma. Alternatively just specify a single device name.

.EXAMPLE	
Get-mPVSdevice -DeviceMac 00-EE-23-43-F5-DD,00-EE-23-43-F5-EE

List devices by a Mac Address Search with their attributes separated with a comma. Alternatively just specify a single mac.

.EXAMPLE	
Get-mPVSdevice -ImageName Windows2003ImageV4 -StoreName Store1 -SiteName Site1

Search for all devices with a specified image name.

.EXAMPLE	
Get-mPVSdevice -CollectionName Group1 -SiteName Site1

Shows a list of devices in a certain collection.

.EXAMPLE	
Get-mPVSdevice -PVSserver PVSserver1

Shows a list of devices that are using a certain PVS server. This also accepts arrays seperated by commas or a list variable

.EXAMPLE	
Get-mPVSdevice | Get-mPVSdeviceExtra

Feed one command in to another to retrieve more information.
	
.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>

    [Cmdletbinding(DefaultParameterSetName='ALL')]
    param(
            [Alias("Name")]
            [Parameter(ParameterSetName='DeviceName',Mandatory=$False,Position=0)]
            $DeviceName,

            [Alias("MacAddress")]
            [Parameter(ParameterSetName='DeviceMac',Mandatory=$False,Position=1)]
            $DeviceMac,

            [Parameter(ParameterSetName='CollectionName',Mandatory=$True,Position=2)]
            $CollectionName,

            [Parameter(ParameterSetName='ImageName',Mandatory=$True,Position=2)]
            [Parameter(ParameterSetName='CollectionName',Mandatory=$True,Position=3)]
            [string]$SiteName,

            [Parameter(ParameterSetName='ImageName',Mandatory=$True,Position=1)]
            [string]$StoreName,

	        [Alias("DiskLocatorName")]
            [Parameter(ParameterSetName='ImageName',Mandatory=$True,Position=0)]
            $ImageName,

            [Parameter(ParameterSetName='PVSserver',Mandatory=$False,Position=0)]
            $PVSserver

    )

    BEGIN{
    
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}
        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli
    
    }#BEGINBLOCKEND

    PROCESS{

        $result = @{}

        if ( $DeviceName -ne $null ){
            $servers = @()
            foreach ($d in $DeviceName){
                $servers += mcli-get device -p DeviceName="$d" -erroraction stop
            }#foreach device
        }#devicename not null
        elseif ($devicemac -ne $null){
            $servers = @()
            foreach ($dm in $DeviceMac){
                $servers += mcli-get device -p DeviceMac="$dm" -erroraction stop
            }#foreachdevicemac
        }

        elseif ( $PVSserver -ne $null){
            $servers = @()
            foreach ($p in $PVSserver){
                $servers += mcli-get device -p ServerName="$p"
            }#pvsserver

        }
        
        elseif( ($ImageName -ne $null) -and ($StoreName -ne $null) ){
            
            $servers = @()
            $servers = Mcli-Get device -p disklocatorname="$ImageName",sitename="$SiteName",storename="$StoreName" -erroraction stop

        }#collectionname not null, Sitename not null

        elseif( ($CollectionName -ne $null) -and ($SiteName -ne $null) ){
            
            $servers = @()
            $servers = mcli-get device -p CollectionName="$CollectionName", SiteName="$SiteName" -erroraction stop
        }#collectionname not null, Sitename not null
        else{
            $servers = mcli-get device
        }#ifdevicename is null

            $servers = $servers | select -Skip 3 | where-object {$_ -ne ""}
            $counting = $servers.count / 32
            $c = 0

            do {
                $server = $servers | select -First 32
                foreach ($line in $server){
                    if ( $line -like "deviceName:*"){
                        $devicename = ($line).substring(12)
                    }#devicename

                    if ( $line -like "deviceid:*"){
                        $deviceid = ($line).substring(10)
                    }#deviceid

                    if ( $line -like "domainControllerName:*"){
                        $domainControllerName = ($line).substring(22)
                    }#domaincontroller
        
                    if ( $line -like "collectionId:*"){
                        $collectionId = ($line).substring(14)
                    }#collectionID

                    if ( $line -like "collectionName:*"){
                        $collectionName = ($line).substring(16)
                    }#collectionName
        
                    if ( $line -like "siteId:*"){
                        $siteId = ($line).substring(8)
                    }#siteId
        
                    if ( $line -like "siteName:*"){
                        $SiteName = ($line).substring(10)
                    }#siteName

        
                    if ( $line -like "description:*"){
                        $Description = ($line).substring(13)
                    }#description


                    if ( $line -like "deviceMac:*"){
                        $DeviceMac = ($line).substring(11)
                    }#devicemac
                
                    if ( $line -like "bootFrom:*"){
                        $BootFrom = ($line).substring(10)
                    }#bootfrom
        
                    if ( $line -like "className:*"){
                        $ClassName = ($line).substring(11)
                    }#classname

                    if ( $line -like "port:*"){
                        $Port = ($line).substring(6)
                    }#port

                    if ( $line -like "enabled:*"){
                        $Enabled = ($line).substring(9)
                    }#Enabled

                    if ( $line -like "localDiskEnabled:*"){
                        $LocalDiskEnabled = ($line).substring(18)
                    }#LocalDiskEnabled

                    if ( $line -like "role:*"){
                        $Role = ($line).substring(6)
                    }#Role

                    if ( $line -like "authentication:*"){
                        $Authentication = ($line).substring(16)
                    }#Authentication

                    if ( $line -like "user:*"){
                        $User = ($line).substring(6)
                    }#User

                    if ( $line -like "password:*"){
                        $Password = ($line).substring(10)
                    }#User

                    if ( $line -like "active:*"){
                        $Active = ($line).substring(8)
                    }#Active

                    if ( $line -like "adTimestamp:*"){
                        $adTimestamp = ($line).substring(13)
                    }#adTimeStamp

                    if ( $line -like "template:*"){
                        $Template = ($line).substring(10)
                    }#Template

                    if ( $line -like "adSignature:*"){
                        $adSignature = ($line).substring(13)
                    }#adSignature

                    if ( $line -like "logLevel:*"){
                        $logLevel = ($line).substring(10)
                    }#logLevel
            
                    if ( $line -like "domainName:*"){
                        $domainName = ($line).substring(12)
                    }#domainName
            
                    if ( $line -like "domainObjectSID:*"){
                        $domainObjectSID = ($line).substring(17)
                    }#domainObjectSID
            
                    if ( $line -like "domainTimeCreated:*"){
                        $domainTimeCreated = ($line).substring(19)
                    }#domainTimeCreated
            
                    if ( $line -like "type:*"){
                        $Type= ($line).substring(6)
                    }#type

                    if ( $line -like "pvdDriveLetter:*"){
                        $pvdDriveLetter= ($line).substring(16)
                    }#pvdDriveLetter
            
                    if ( $line -like "localWriteCacheDiskSize:*"){
                        $localWriteCacheDiskSize= ($line).substring(25)
                    }#localWriteCacheDiskSize
            
                    if ( $line -like "bdmBoot:*"){
                        $bdmBoot= ($line).substring(9)
                    }#bdmBoot
            
                    if ( $line -like "virtualHostingPoolId:*"){
                        $virtualHostingPoolId= ($line).substring(22)
                    }#virtualHostingPoolId

                }#foreachend
              
                $result = @{
                             'DeviceName' = $devicename;
                             'DeviceID' = $deviceid;
                             'DomainControllerName' = $domainControllerName;
                             'CollectionID' = $collectionId;
                             'CollectionName' = $collectionName;
                             'SiteID' = $siteId;
                             'SiteName' = $SiteName;
                             'Description' = $Description;
                             'DeviceMac' = $DeviceMac;
                             'BootFrom' = $BootFrom;
                             'ClassName' = $ClassName;
                             'Port' = $Port;
                             'Enabled' = $Enabled;
                             'LocalDiskEnabled' = $LocalDiskEnabled;
                             'Role' = $Role;
                             'Authentication' = $Authentication;
                             'User' = $User;
                             'Password' = $Password;
                             'Active' = $Active;
                             'adTimestamp' = $adTimestamp;
                             'Template' = $Template;
                             'adSignature' = $adSignature;
                             'logLevel' = $logLevel;
                             'domainName' = $domainName;
                             'domainObjectSID' = $domainObjectSID;
                             'domainTimeCreated' = $domainTimeCreated;
                             'Type' = $Type;
                             'pvdDriveLetter' = $pvdDriveLetter;
                             'localWriteCacheDiskSize' = $localWriteCacheDiskSize;
                             'bdmBoot' = $bdmBoot;
                             'virtualHostingPoolId' = $virtualHostingPoolId
                 }#result
                 $obj = New-Object -TypeName PSObject -Property $result
		
		 $obj.psobject.typenames.insert(0,'MorNando.PVS.GetDevice')
		 Write-Output $obj
        
                 $servers = $servers | select -skip 32
                 $c++

            } until($c -ge $counting)

    }#PROCESS END

    END{
        try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}
    }

}#End of Get-mPVSdevice

Function Get-mPVSdeviceExtra {
    
<#
.SYNOPSIS
Lists citrix PVS devices and their attributes by DeviceName search. It's similar to the Get-mPVSdevice cmdlet.                    

Get-mPVSdeviceExtra is slower than Get-mPVSdevice and can only search by name. However, it provides more information. 

One of the main benefits is that you can pipe the results from Get-mPVSdevice in to Get-mPVSdeviceExtra.

Feed one command in to another to retrieve more information.                              
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly.                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.

.PARAMETER DeviceName
Provide the name of the device you which to search on. This parameter can only be used by itself but it can accept arrays.

.EXAMPLE	
Get-mPVSdevice | Get-mPVSdeviceExtra

Feed one command in to another to retrieve more information.

.EXAMPLE	
Get-mPVSdeviceExtra -DeviceName device1,device2,device3
	
List devices by a name search with their attributes by seperating with a comma. Alternatively just specify a single device name.

.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>

    [Cmdletbinding(DefaultParameterSetName='DeviceName')]
    param(
            [Alias("Name")]
            [Parameter(ParameterSetName='DeviceName',Mandatory=$True,Position=0,ValueFromPipelinebyPropertyName=$True)]
            $DeviceName

    )

    BEGIN{
        
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}

        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli

    }#BEGINBLOCKFINISHED

    PROCESS{
            $servers = @()
            foreach ($d in $DeviceName){

                $servers += Mcli-Get deviceinfo -p devicename="$d" -erroraction stop | select -Skip 3 | where-object {$_ -ne ""}

            }#foreach device

            $counting = $servers.count / 47
            $c = 0

            do {
                $server = $servers | select -First 47
                foreach ($line in $server){
                    if ( $line -like "deviceName:*"){
                        $devicename = ($line).substring(12)
                    }#devicename

                    if ( $line -like "deviceid:*"){
                        $deviceid = ($line).substring(10)
                    }#deviceid

                    if ( $line -like "domainControllerName:*"){
                        $domainControllerName = ($line).substring(22)
                    }#domaincontroller
        
                    if ( $line -like "collectionId:*"){
                        $collectionId = ($line).substring(14)
                    }#collectionID

                    if ( $line -like "collectionName:*"){
                        $collectionName = ($line).substring(16)
                    }#collectionName
        
                    if ( $line -like "siteId:*"){
                        $siteId = ($line).substring(8)
                    }#siteId
        
                    if ( $line -like "siteName:*"){
                        $SiteName = ($line).substring(10)
                    }#siteName

        
                    if ( $line -like "description:*"){
                        $Description = ($line).substring(13)
                    }#description


                    if ( $line -like "deviceMac:*"){
                        $DeviceMac = ($line).substring(11)
                    }#devicemac
                
                    if ( $line -like "bootFrom:*"){
                        $BootFrom = ($line).substring(10)
                    }#bootfrom
        
                    if ( $line -like "className:*"){
                        $ClassName = ($line).substring(11)
                    }#classname

                    if ( $line -like "port:*"){
                        $Port = ($line).substring(6)
                    }#port

                    if ( $line -like "enabled:*"){
                        $Enabled = ($line).substring(9)
                    }#Enabled

                    if ( $line -like "localDiskEnabled:*"){
                        $LocalDiskEnabled = ($line).substring(18)
                    }#LocalDiskEnabled

                    if ( $line -like "role:*"){
                        $Role = ($line).substring(6)
                    }#Role

                    if ( $line -like "authentication:*"){
                        $Authentication = ($line).substring(16)
                    }#Authentication

                    if ( $line -like "user:*"){
                        $User = ($line).substring(6)
                    }#User

                    if ( $line -like "password:*"){
                        $Password = ($line).substring(10)
                    }#User

                    if ( $line -like "active:*"){
                        $Active = ($line).substring(8)
                    }#Active

                    if ( $line -like "adTimestamp:*"){
                        $adTimestamp = ($line).substring(13)
                    }#adTimeStamp

                    if ( $line -like "template:*"){
                        $Template = ($line).substring(10)
                    }#Template

                    if ( $line -like "adSignature:*"){
                        $adSignature = ($line).substring(13)
                    }#adSignature

                    if ( $line -like "logLevel:*"){
                        $logLevel = ($line).substring(10)
                    }#logLevel
            
                    if ( $line -like "domainName:*"){
                        $domainName = ($line).substring(12)
                    }#domainName
            
                    if ( $line -like "domainObjectSID:*"){
                        $domainObjectSID = ($line).substring(17)
                    }#domainObjectSID
            
                    if ( $line -like "domainTimeCreated:*"){
                        $domainTimeCreated = ($line).substring(19)
                    }#domainTimeCreated
            
                    if ( $line -like "type:*"){
                        $Type= ($line).substring(6)
                    }#type

                    if ( $line -like "pvdDriveLetter:*"){
                        $pvdDriveLetter= ($line).substring(16)
                    }#pvdDriveLetter
            
                    if ( $line -like "localWriteCacheDiskSize:*"){
                        $localWriteCacheDiskSize = ($line).substring(25)
                    }#localWriteCacheDiskSize
            
                    if ( $line -like "bdmBoot:*"){
                        $bdmBoot = ($line).substring(9)
                    }#bdmBoot
            
                    if ( $line -like "virtualHostingPoolId:*"){
                        $virtualHostingPoolId = ($line).substring(22)
                    }#virtualHostingPoolId

                    if ( $line -like "ip:*"){
                        $ipaddress = ($line).substring(4)
                    }#IP
                    
                    if ( $line -like "serverPortConnection:*"){
                        $serverPortConnection = ($line).substring(22)
                    }#serverPortConnection
                    
                    if ( $line -like "serverIpConnection:*"){
                        $serverIpConnection = ($line).substring(20)
                    }#serverIpConnection
                    
                    if ( $line -like "serverId:*"){
                        $serverId = ($line).substring(10)
                    }#serverId

                    if ( $line -like "serverName:*"){
                        $serverName = ($line).substring(12)
                    }#serverName

                    if ( $line -like "diskLocatorId:*"){
                        $diskLocatorId = ($line).substring(15)
                    }#diskLocatorId
                    
                    if ( $line -like "diskLocatorName:*"){
                        $diskLocatorName = ($line).substring(17)
			$imageName = ($line).substring(17)
			$imagename = $imagename.split("{\}") | select -last 1
                    }#diskLocatorName
                    
                    if ( $line -like "diskVersion:*"){
                        $diskVersion = ($line).substring(13)
                    }#diskLocatorName
                    
                    if ( $line -like "diskVersionAccess:*"){
                        $diskVersionAccess = ($line).substring(19)
                    }#diskLocatorName
                    
                    if ( $line -like "diskFileName:*"){
                        $diskFileName = ($line).substring(14)
                    }#diskFileName
                    
                    if ( $line -like "status:*"){
                        $status = ($line).substring(8)
                    }#status
                    
                    if ( $line -like "licenseType:*"){
                        $licenseType = ($line).substring(13)
                    }#licensetype
                    
                    if ( $line -like "makLicenseActivated:*"){
                        $makLicenseActivated = ($line).substring(21)
                    }#makLicenseActivated
                    
                    if ( $line -like "model:*"){
                        $model = ($line).substring(7)
                    }#model
                    
                    if ( $line -like "license:*"){
                        $license = ($line).substring(9)
                    }#license

                }#foreachend
              
                $result = @{
                             'DeviceName' = $devicename;
                             'DeviceID' = $deviceid;
                             'DomainControllerName' = $domainControllerName;
                             'CollectionID' = $collectionId;
                             'CollectionName' = $collectionName;
                             'SiteID' = $siteId;
                             'SiteName' = $SiteName;
                             'Description' = $Description;
                             'DeviceMac' = $DeviceMac;
                             'BootFrom' = $BootFrom;
                             'ClassName' = $ClassName;
                             'Port' = $Port;
                             'Enabled' = $Enabled;
                             'LocalDiskEnabled' = $LocalDiskEnabled;
                             'Role' = $Role;
                             'Authentication' = $Authentication;
                             'User' = $User;
                             'Password' = $Password;
                             'Active' = $Active;
                             'adTimestamp' = $adTimestamp;
                             'Template' = $Template;
                             'adSignature' = $adSignature;
                             'logLevel' = $logLevel;
                             'domainName' = $domainName;
                             'domainObjectSID' = $domainObjectSID;
                             'domainTimeCreated' = $domainTimeCreated;
                             'Type' = $Type;
                             'pvdDriveLetter' = $pvdDriveLetter;
                             'localWriteCacheDiskSize' = $localWriteCacheDiskSize;
                             'bdmBoot' = $bdmBoot;
                             'virtualHostingPoolId' = $virtualHostingPoolId;
                             'IPAddress' = $ipaddress;
                             'ServerPortConnection' = $serverPortConnection;
                             'ServerIpConnection' = $serverIpConnection;
                             'ServerId' = $serverId;
                             'ServerName' = $serverName;
                             'diskLocatorId' = $diskLocatorId;
                             'diskLocatorName' = $diskLocatorName;
                             'diskVersion' = $diskVersion;
                             'diskVersionAccess' = $diskVersionAccess;
                             'diskFileName' = $diskFileName;
                             'Status' = $status;
                             'licenseType' = $licenseType;
                             'makLicenseActivated' = $makLicenseActivated;
                             'Model' = $model;
                             'License' = $license;
			     'ImageName' = $imagename                        
                 }#result
                 $obj = New-Object -TypeName PSObject -Property $result
		 $obj.psobject.typenames.insert(0,'MorNando.PVS.GetDeviceExtra')
		 Write-Output $obj
        
                 $servers = $servers | select -skip 47
                 $c++

            } until($c -ge $counting)
    
    }#PROCESS end

    END{
	        try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}

}

}#Function end

Function Get-mPVSdiskInfo {
<#
.SYNOPSIS
Lists citrix PVS Disks and their attributes.                                                   
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly.                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.

.EXAMPLE	
Get-mPVSdiskInfo
	
No parameter necessary.

.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>
    
    BEGIN{
        
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}

        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli

    }#BEGIN FINISHED

    PROCESS {

        $diskinfo = mcli-get diskinfo -erroraction stop | select -Skip 3 | where-object {$_ -ne ""}

        $counting = $diskinfo.count / 47

        $c = 0

            do {
                $disk = $diskinfo | select -First 47
                foreach ($line in $disk){
                    if ( $line -like "diskLocatorId:*"){
                        $diskLocatorId = ($line).substring(15)
                    }#diskLocatorId

                    if ( $line -like "diskLocatorName:*"){
                        $diskLocatorName = ($line).substring(17)
                    }#diskLocatorName

                    if ( $line -like "siteId:*"){
                        $siteId = ($line).substring(8)
                    }#siteId
        
                    if ( $line -like "siteName:*"){
                        $siteName = ($line).substring(10)
                    }#siteName

                    if ( $line -like "storeId:*"){
                        $storeId = ($line).substring(9)
                    }#storeId
        
                    if ( $line -like "siteid:*"){
                        $siteId = ($line).substring(8)
                    }#siteId
        
                    if ( $line -like "description:*"){
                        $description = (($line).substring(13))

                    }#description

                    if ( $line -like "menuText:*"){
                        $menuText = ($line).substring(10)
                    }#menuText


                    if ( $line -like "serverId:*"){
                        $serverId = ($line).substring(10)
                    }#serverId
                
                    if ( $line -like "serverName:*"){
                        $serverName = ($line).substring(12)
                    }#serverName
        
                    if ( $line -like "enabled:*"){
                        $enabled = ($line).substring(9)
                    }#enabled

                    if ( $line -like "role:*"){
                        $role = ($line).substring(6)
                    }#role

                    if ( $line -like "mapped:*"){
                        $mapped = ($line).substring(8)
                    }#mapped

                    if ( $line -like "active:*"){
                        $active = ($line).substring(8)
                    }#active

                    if ( $line -like "rebalanceEnabled:*"){
                        $rebalanceEnabled = ($line).substring(18)
                    }#rebalanceEnabled

                    if ( $line -like "rebalanceTriggerPercent:*"){
                        $rebalanceTriggerPercent = ($line).substring(25)
                    }#rebalanceTriggerPercent

                    if ( $line -like "subnetAffinity:*"){
                        $subnetAffinity = ($line).substring(16)
                    }#subnetAffinity

                    if ( $line -like "diskUpdateDeviceId:*"){
                        $diskUpdateDeviceId = ($line).substring(20)
                    }#diskUpdateDeviceId

                    if ( $line -like "diskUpdateDeviceName:*"){
                        $diskUpdateDeviceName = ($line).substring(22)
                    }#diskUpdateDeviceName

                    if ( $line -like "class:*"){
                        $class = ($line).substring(7)
                    }#class

                    if ( $line -like "imageType:*"){
                        $imageType = ($line).substring(11)
                    }#imageType

                    if ( $line -like "diskSize:*"){
                        $diskSize = ($line).substring(10)
                    }#diskSize

                    if ( $line -like "vhdBlockSize:*"){
                        $vhdBlockSize = ($line).substring(14)
                    }#vhdBlockSize
            
                    if ( $line -like "writeCacheSize:*"){
                        $writeCacheSize = ($line).substring(16)
                    }#writeCacheSize
            
                    if ( $line -like "autoUpdateEnabled:*"){
                        $autoUpdateEnabled = ($line).substring(19)
                    }#autoUpdateEnabled
            
                    if ( $line -like "activationDateEnabled:*"){
                        $activationDateEnabled = ($line).substring(23)
                    }#activationDateEnabled
            
                    if ( $line -like "adPasswordEnabled:*"){
                        $adPasswordEnabled= ($line).substring(19)
                    }#adPasswordEnabled

                    if ( $line -like "haEnabled:*"){
                        $haEnabled= ($line).substring(11)
                    }#haEnabled
            
                    if ( $line -like "printerManagementEnabled:*"){
                        $printerManagementEnabled = ($line).substring(26)
                    }#printerManagementEnabled
            
                    if ( $line -like "writeCacheType:*"){
                        $writeCacheType = ($line).substring(16)
                    }#writeCacheType
            
                    if ( $line -like "licenseMode:*"){
                        $licenseMode = ($line).substring(13)
                    }#licenseMode

                    if ( $line -like "activeDate:*"){
                        $activeDate = ($line).substring(12)
                    }#activeDate
                    
                    if ( $line -like "longDescription:*"){
                        $longDescription = ($line).substring(17)

                    }#longDescription
                    
                    if ( $line -like "serialNumber:*"){
                        $serialNumber = ($line).substring(14)
                    }#serialNumber
                    
                    if ( $line -like "date:*"){
                        $date = ($line).substring(6)
                    }#date

                    if ( $line -like "author:*"){
                        $author = ($line).substring(8)
                    }#author

                    if ( $line -like "title:*"){
                        $title = ($line).substring(7)
                    }#title

                    if ( $line -like "company:*"){
                        $company = ($line).substring(9)
                    }#company
                    
                    if ( $line -like "internalName:*"){
                        $internalName = ($line).substring(14)
                    }#internalName
                    
                    if ( $line -like "originalFile:*"){
                        $originalFile = ($line).substring(14)
                    }#originalFile
                    
                    if ( $line -like "hardwareTarget:*"){
                        $hardwareTarget = ($line).substring(16)
                    }#hardwareTarget
                    
                    if ( $line -like "majorRelease:*"){
                        $majorRelease = ($line).substring(14)
                    }#majorRelease
                    
                    if ( $line -like "minorRelease:*"){
                        $minorRelease = ($line).substring(14)
                    }#minorRelease
                    
                    if ( $line -like "build:*"){
                        $build = ($line).substring(7)
                    }#build
                    
                    if ( $line -like "deviceCount:*"){
                        $deviceCount = ($line).substring(13)
                    }#deviceCount
                    
                    if ( $line -like "locked:*"){
                        $locked= ($line).substring(8)
                    }#locked
                    
                }#foreachend
              
                $result = @{
                             'DiskLocatorId' = $diskLocatorId;
                             'DiskLocatorName' = $diskLocatorName;
                             'SiteId' = $siteId;
                             'SiteName' = $siteName;
                             'StoreId' = $storeId;
                             'StoreName' = $storeName;
			     'DeviceCount' = $deviceCount;
                             'Description' = $Description;
                             'MenuText' = $menuText;
                             'ServerId' = $serverId;
                             'ServerName' = $serverName;
                             'Enabled' = $enabled;
                             'Role' = $role;
                             'Mapped' = $Mapped;
                             'Active' = $Active;
                             'RebalanceEnabled' = $RebalanceEnabled;
                             'RebalanceTriggerPercent' = $rebalanceTriggerPercent;
                             'SubnetAffinity' = $subnetAffinity;
                             'DiskUpdateDeviceId' = $diskUpdateDeviceId;
                             'DiskUpdateDeviceName' = $diskUpdateDeviceName;
                             'Class' = $class;
                             'imageType' = $imageType;
                             'DiskSize' = $diskSize;
                             'VhdBlockSize' = $VhdBlockSize;
                             'WriteCacheSize' = $writeCacheSize;
                             'AutoUpdateEnabled' = $autoUpdateEnabled;
                             'ActivationDateEnabled' = $activationDateEnabled;
                             'AdPasswordEnabled' = $adPasswordEnabled;
                             'haEnabled' = $haEnabled;
                             'PrinterManagementEnabled' = $printerManagementEnabled;
                             'WriteCacheType' = $writeCacheType;
                             'LicenseMode' = $licenseMode;
                             'ActiveDate' = $activeDate;
                             'LongDescription' = $longDescription;
                             'SerialNumber' = $SerialNumber;
                             'Date' = $date;
                             'Author' = $author;
                             'Title' = $title;
                             'Company' = $Company;
                             'InternalName' = $InternalName;
                             'OriginalFile' = $originalFile;
                             'HardwareTarget' = $hardwareTarget;
                             'MajorRelease' = $majorRelease;
                             'MinorRelease' = $minorRelease;
                             'Build' = $build;
                             'Locked' = $locked                         
                 }#result

                 $obj = New-Object -TypeName PSObject -Property $result
		 $obj.psobject.typenames.insert(0,'MorNando.PVS.GetDiskInfo')
		 Write-Output $obj
	         
		 $diskinfo = $diskinfo | select -skip 47
                 $c++

            } until($c -ge $counting)
    
    }#PROCESS end

    END{
	
	    try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}
}
}#end of getdiskinfo

Function Get-mPVSserverInfo {

<#
.SYNOPSIS
Displays in depth PVS server information.                               
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly. We need object data!                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.

.PARAMETER PVSserver
Provide the name of the device you which to search on. This parameter can only be used by itself but it can accept arrays.
	
.EXAMPLE	
Get-mPVSserverInfo -PVSServer Server1
	
This will list all devices with their attributes
	
.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>
[Cmdletbinding(DefaultParameterSetName='PVSServer')]
    param(
            [Alias("ServerName")]
            [Alias("Name")]
            [Parameter(ParameterSetName='PVSServer',Mandatory=$False,Position=0)]
            [string[]]$PVSserver
        )

     BEGIN{
        
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}

        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli

    }#BEGIN FINISHED

    PROCESS{

        if ($PVSserver -ne $null){
            $servers = @()
            foreach ($s in $PVSserver){
                $servers += mcli-get serverinfo -p servername="$s" | select -skip 3 | where-object { $_ -ne "" }
            }#foreach servers end

        }#if pvsservers not null
        else {

            $servers = mcli-get serverinfo | select -skip 3 | where-object { $_ -ne "" }

        }

         $counting = $servers.count / 48
                    $c = 0

                        do {
                            $serversingle = $servers | select -First 48

                            foreach ($line in $serversingle){

                                if ( $line -like "serverId:*"){
                                    $ServerID = ($line).substring(10)
                                }#serverId

                                if ( $line -like "serverName:*"){
                                    $ServerNameprop = ($line).substring(12)
                                }#serverName

                                if ( $line -like "siteId:*"){
                                    $siteId = ($line).substring(8)
                                }#siteId
        
                                 if ( $line -like "siteName:*"){
                                    $siteName = ($line).substring(10)
                                }#siteName

                                if ( $line -like "description:*"){
                                    $Description = ($line).substring(13)
                                }#description
                    
                                if ( $line -like "adMaxPasswordAge:*"){
                                    $adMaxPasswordAge = ($line).substring(18)
                                }#adMaxPasswordAge

                                if ( $line -like "licenseTimeout:*"){
                                    $licenseTimeout = ($line).substring(16)
                                }#licenseTimeout

                                if ( $line -like "vDiskCreatePacing:*"){
                                    $vDiskCreatePacing = ($line).substring(19)
                                }#vDiskCreatePacing

                                if ( $line -like "firstPort:*"){
                                    $firstPort = ($line).substring(11)
                                }#firstPort

                                if ( $line -like "lastPort:*"){
                                    $lastPort = ($line).substring(10)
                                }#lastPort

                                if ( $line -like "threadsPerPort:*"){
                                    $threadsPerPort = ($line).substring(16)
                                }#threadsPerPort

                                if ( $line -like "buffersPerThread:*"){
                                    $buffersPerThread = ($line).substring(18)
                                }#buffersPerThread

                                if ( $line -like "serverCacheTimeout:*"){
                                    $serverCacheTimeout = ($line).substring(20)
                                }#serverCacheTimeout

                                if ( $line -like "ioBurstSize:*"){
                                    $ioBurstSize = ($line).substring(13)
                                }#ioBurstSize

                                if ( $line -like "maxTransmissionUnits:*"){
                                    $maxTransmissionUnits = ($line).substring(22)
                                }#maxTransmissionUnits

                                if ( $line -like "maxBootDevicesAllowed:*"){
                                    $maxBootDevicesAllowed = ($line).substring(23)
                                }#maxBootDevicesAllowed

                                if ( $line -like "maxBootSeconds:*"){
                                    $maxBootSeconds = ($line).substring(16)
                                }#maxBootSeconds

                                if ( $line -like "bootPauseSeconds:*"){
                                    $bootPauseSeconds = ($line).substring(18)
                                }#bootPauseSeconds

                                if ( $line -like "adMaxPasswordAgeEnabled:*"){
                                    $adMaxPasswordAgeEnabled = ($line).substring(25)
                                }#adMaxPasswordAgeEnabled

                                if ( $line -like "eventLoggingEnabled:*"){
                                    $eventLoggingEnabled = ($line).substring(21)
                                }#eventLoggingEnabled

                                if ( $line -like "nonBlockingIoEnabled:*"){
                                    $nonBlockingIoEnabled = ($line).substring(22)
                                }#nonBlockingIoEnabled

                                if ( $line -like "role:*"){
                                    $role = ($line).substring(6)
                                }#role

                                if ( $line -like "ip:*"){
                                    $ipaddress = ($line).substring(4)
                                }#ip

                                if ( $line -like "initialQueryConnectionPoolSize:*"){
                                    $initialQueryConnectionPoolSize = ($line).substring(32)
                                }#initialQueryConnectionPoolSize

                                if ( $line -like "initialTransactionConnectionPoolSize:*"){
                                    $initialTransactionConnectionPoolSize = ($line).substring(38)
                                }#initialTransactionConnectionPoolSize

                                if ( $line -like "maxQueryConnectionPoolSize:*"){
                                    $maxQueryConnectionPoolSize = ($line).substring(28)
                                }#maxQueryConnectionPoolSize

                                if ( $line -like "maxTransactionConnectionPoolSize:*"){
                                    $maxTransactionConnectionPoolSize = ($line).substring(34)
                                }#maxTransactionConnectionPoolSize

                                if ( $line -like "refreshInterval:*"){
                                    $refreshInterval = ($line).substring(17)
                                }#refreshInterval

                                if ( $line -like "unusedDbConnectionTimeout:*"){
                                    $unusedDbConnectionTimeout = ($line).substring(27)
                                }#unusedDbConnectionTimeout

                                if ( $line -like "busyDbConnectionRetryCount:*"){
                                    $busyDbConnectionRetryCount = ($line).substring(28)
                                }#busyDbConnectionRetryCount

                                if ( $line -like "busyDbConnectionRetryInterval:*"){
                                    $busyDbConnectionRetryInterval = ($line).substring(31)
                                }#busyDbConnectionRetryInterval

                                if ( $line -like "localConcurrentIoLimit:*"){
                                    $localConcurrentIoLimit = ($line).substring(24)
                                }#localConcurrentIoLimit

                                if ( $line -like "remoteConcurrentIoLimit:*"){
                                    $remoteConcurrentIoLimit = ($line).substring(25)
                                }#remoteConcurrentIoLimit

                                if ( $line -like "ramDiskIpAddress:*"){
                                    $ramDiskIpAddress = ($line).substring(18)
                                }#ramDiskIpAddress

                                if ( $line -like "ramDiskTimeToLive:*"){
                                    $ramDiskTimeToLive = ($line).substring(19)
                                }#ramDiskTimeToLive

                                if ( $line -like "ramDiskInvitationType:*"){
                                    $ramDiskInvitationType = ($line).substring(23)
                                }#ramDiskInvitationType

                                if ( $line -like "ramDiskInvitationPeriod:*"){
                                    $ramDiskInvitationPeriod = ($line).substring(25)
                                }#ramDiskInvitationPeriod

                                if ( $line -like "active:*"){
                                    $active = ($line).substring(8)
                                }#active

                                if ( $line -like "logLevel:*"){
                                    $logLevel = ($line).substring(10)
                                }#logLevel

                                if ( $line -like "logFileSizeMax:*"){
                                    $logFileSizeMax = ($line).substring(16)
                                }#logFileSizeMax

                                if ( $line -like "logFileBackupCopiesMax:*"){
                                    $logFileBackupCopiesMax = ($line).substring(24)
                                }#logFileBackupCopiesMax

                                if ( $line -like "powerRating:*"){
                                    $powerRating = ($line).substring(13)
                                }#powerRating

                                if ( $line -like "serverFqdn:*"){
                                    $serverFqdn = ($line).substring(12)
                                }#serverFqdn

                                if ( $line -like "managementIp:*"){
                                    $managementIp = ($line).substring(14)
                                }#managementIp

                                if ( $line -like "contactIp:*"){
                                    $contactIp = ($line).substring(11)
                                }#contactIp

                                if ( $line -like "contactPort:*"){
                                    $contactPort = ($line).substring(13)
                                }#contactPort
                                                                                  
                            }#foreachend
              
                            $result = @{
                                         'ServerID' = $ServerID;
                                         'ServerName' = $ServerNameprop;
                                         'SiteID' = $siteId;
                                         'SiteName' = $siteName;
                                         'Description' = $description;
                                         'adMaxPasswordAge' = $adMaxPasswordAge;
                                         'LicenseTimeout' = $licenseTimeout;
                                         'vDiskCreatePacing' = $vDiskCreatePacing;
                                         'FirstPort' = $firstPort;
                                         'lastPort' = $lastPort;
                                         'threadsPerPort' = $threadsPerPort;
                                         'buffersPerThread' = $buffersPerThread;
                                         'serverCacheTimeout' = $serverCacheTimeout;
                                         'ioBurstSize' = $ioBurstSize;
                                         'maxTransmissionUnits' = $maxTransmissionUnits;
                                         'maxBootDevicesAllowed' = $maxBootDevicesAllowed;
                                         'maxBootSeconds' = $maxBootSeconds;
                                         'bootPauseSeconds' = $bootPauseSeconds;
                                         'adMaxPasswordAgeEnabled' = $adMaxPasswordAgeEnabled;
                                         'eventLoggingEnabled' = $eventLoggingEnabled; 
                                         'nonBlockingIoEnabled' = $nonBlockingIoEnabled;
                                         'role' = $role;
                                         'IPAddress' = $IPAddress;
                                         'initialQueryConnectionPoolSize' = $initialQueryConnectionPoolSize;
                                         'initialTransactionConnectionPoolSize' = $initialTransactionConnectionPoolSize;
                                         'maxQueryConnectionPoolSize' = $maxQueryConnectionPoolSize;
                                         'maxTransactionConnectionPoolSize' = $maxTransactionConnectionPoolSize;
                                         'refreshInterval' = $refreshInterval;
                                         'unusedDbConnectionTimeout' = $unusedDbConnectionTimeout;
                                         'busyDbConnectionRetryCount' = $busyDbConnectionRetryCount; 
                                         'busyDbConnectionRetryInterval' = $busyDbConnectionRetryInterval;
                                         'localConcurrentIoLimit' = $localConcurrentIoLimit;
                                         'remoteConcurrentIoLimit' = $remoteConcurrentIoLimit;
                                         'ramDiskInvitationType' = $ramDiskInvitationType;
                                         'ramDiskInvitationPeriod' = $ramDiskInvitationPeriod;
                                         'active' = $active;
                                         'logLevel' = $logLevel;
                                         'logFileSizeMax' = $logFileSizeMax;
                                         'logFileBackupCopiesMax' = $logFileBackupCopiesMax;
                                         'powerRating' = $powerRating; 
                                         'serverFqdn' = $serverFqdn;
                                         'managementIp' = $managementIp;
                                         'contactIp' = $contactIp;
                                         'contactPort' = $contactPort
                                                                                                                                                              
                             }#result

                             $obj = New-Object -TypeName PSObject -Property $result
		             $obj.psobject.typenames.insert(0,'MorNando.PVS.GetPVSserverInfo')
		             Write-Output $obj
        
                             $servers = $servers | select -skip 48
                             $c++

                        } until($c -ge $counting)
    }#processfinished

    END{
        try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}
    }

}#end of Get-mPVSserverInfo

Function Get-mPVSserverStatus {
<#
.SYNOPSIS
Displays the status of the named PVS Server.                               
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly. We need object data!                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.

.PARAMETER PVSserver
Provide the name of the device you which to search on. This parameter can only be used by itself but it can accept arrays.
	
.EXAMPLE	
Get-mPVSserverStatus -PVSServer Server1
	
This will list a devices status and a few extra attributes.

Get-mPVSserverStatus -PVSServer Server1,Server2
	
This will list a devices status and a few extra attributes.It can also accept a variable list or multiple servers seperated by a comma.
	
.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>
    [Cmdletbinding(DefaultParameterSetName='All')]
    param(
            [Alias("ServerName")]
            [Alias("Name")]
            [Parameter(ParameterSetName='PVSServer',Mandatory=$False,Position=0)]
            $PVSserver
        )

     BEGIN{
        
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}

        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli

    }#BEGIN FINISHED

    PROCESS{
        
        if ($PVSserver -ne $null){

            $status = @()
            foreach ($s in $PVSserver){
                $status += Mcli-Get serverstatus -p servername="$s" | select -skip 3 | where-object { $_ -ne "" }
            }

        }#if pvsserver isnt null

        else{
            $status = @()
            $serverinf = Get-mPVSserverInfo | select -ExpandProperty "ServerName"
            Add-PSSnapin mclipssnapin -ErrorAction SilentlyContinue
            foreach ($s in $serverinf) {

                $status += mcli-Get serverstatus -p servername="$s" | select -skip 3 | where-object { $_ -ne "" }

            }#foreach s in serverinf

        }#elseend

            $counting = $status.count / 7
            $c = 0

                do {
                    $statussingle = $status | select -First 7

                    foreach ($line in $statussingle){

                        if ( $line -like "serverId:*"){
                            $ServerID = ($line).substring(10)
                        }#serverId

                        if ( $line -like "serverName:*"){
                            $ServerName = ($line).substring(12)
                        }#serverName

                        if ( $line -like "ip:*"){
                            $IPAddress = ($line).substring(4)
                        }#ipaddress
        
                         if ( $line -like "port:*"){
                            $Port = ($line).substring(6)
                        }#port

                        if ( $line -like "deviceCount:*"){
                            $DeviceCount = ($line).substring(13)
                        }#deviceCount 
                    
                        if ( $line -like "status:*"){
                            $Statusprop = ($line).substring(8)
                        }#status
                                                          
                    }#foreachend
              
                    $result = @{
                                 'ServerID' = $ServerID;
                                 'ServerName' = $ServerName;
                                 'IPAddress' = $IPAddress;
                                 'Port' = $Port;
                                 'DeviceCount' = $DeviceCount;
                                 'Status' = $Statusprop                                                   
                     }#result

                     $obj = New-Object -TypeName PSObject -Property $result
		     $obj.psobject.typenames.insert(0,'MorNando.PVS.GetPVSserverStatus')
		     Write-Output $obj
        
                     $status = $status | select -skip 7
                     $c++

                } until($c -ge $counting)

        
    }#PROCESSFINISHED

    END{
        try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}
    }

} #END OF Get-mPVSserverStatus function

Function Get-mPVScollection {
<#
.SYNOPSIS
Displays a list of collection names and their properties.                               
	
This cmdlet was written by Chris Morland. Please contact me on chrismorland@lambtonsolutions.co.uk with any queries, bugs or feature requests.

.DESCRIPTION
This module is a rewritten version of mclipssnapin due to the original cmdlets returning string data and generally not being user friendly. We need object data!                            
	
Therefore this has a dependency on the mclipssnapin. In order to load the mcli snapin on a PVS server a Dll file needs to be registered.                                                                                                    
Open a cmd prompt (AS ADMINSTRATOR) and run:                                                                                                        
"C:\windows\microsoft.net\framework64\v2.0.50727\InstallUtil.exe" "C:\Program Files\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
                                                                                                                                        
Note: Using PowerShell to register the dll file will not work. It needs to be cmd.

MorNandoPVSTools provides a good way to manage your PVS environment using PowerShell.
	
.EXAMPLE	
Get-mPVScollection
	
.NOTES
This module and cmdlet was created by Chris Morland

.LINK
http://www.lambtonsolutions.co.uk
#>

BEGIN{
        
        $mcli = get-pssnapin | where { $_.name -eq "McliPSSnapIn"}

        if ($mcli -eq $null){
	        try{
                Add-PSSnapin *mcli*
	        }
	        catch{}
        }#mcli
}#Begin block finished

PROCESS{

    $collection = mcli-get collection | select -skip 3 | where-object { $_ -ne "" }

    $counting = $collection.count / 19
            $c = 0

                do {
                    $collect = $collection | select -First 19

                    foreach ($line in $collect){

                        if ( $line -like "collectionId:*"){
                            $collectionId = ($line).substring(14)
                        }#collectionId

                        if ( $line -like "collectionName:*"){
                            $collectionName = ($line).substring(16)
                        }#collectionName

                        if ( $line -like "siteId:*"){
                            $siteId = ($line).substring(8)
                        }#siteId
        
                         if ( $line -like "siteName:*"){
                            $siteName = ($line).substring(10)
                        }#siteName

                        if ( $line -like "description:*"){
                            $description = ($line).substring(13)
                        }#description: 
                    
                        if ( $line -like "templateDeviceId:*"){
                            $templateDeviceId = ($line).substring(18)
                        }#templateDeviceId

                        if ( $line -like "templateDeviceName:*"){
                            $templateDeviceName = ($line).substring(20)
                        }#templateDeviceName

                        if ( $line -like "lastAutoAddDeviceNumber:*"){
                            $lastAutoAddDeviceNumber = ($line).substring(25)
                        }#lastAutoAddDeviceNumber

                        if ( $line -like "enabled:*"){
                            $enabled = ($line).substring(9)
                        }#enabled

                        if ( $line -like "deviceCount:*"){
                            $deviceCount = ($line).substring(13)
                        }#deviceCount

                        if ( $line -like "deviceWithPVDCount:*"){
                            $deviceWithPVDCount = ($line).substring(20)
                        }#deviceWithPVDCount

                        if ( $line -like "activeDeviceCount:*"){
                            $activeDeviceCount = ($line).substring(19)
                        }#activeDeviceCount

                        if ( $line -like "makActivateNeededCount:*"){
                            $makActivateNeededCount = ($line).substring(24)
                        }#makActivateNeededCount

                        if ( $line -like "autoAddPrefix:*"){
                            $autoAddPrefix = ($line).substring(15)
                        }#autoAddPrefix

                        if ( $line -like "autoAddSuffix:*"){
                            $autoAddSuffix = ($line).substring(15)
                        }#autoAddSuffix

                        if ( $line -like "autoAddZeroFill:*"){
                            $autoAddZeroFill = ($line).substring(17)
                        }#autoAddZeroFill

                        if ( $line -like "autoAddNumberLength:*"){
                            $autoAddNumberLength = ($line).substring(21)
                        }#autoAddNumberLength

                        if ( $line -like "role:*"){
                            $role = ($line).substring(6)
                        }#role
                                                          
                    }#foreachend
              
                    $result = @{
                                 'collectionId:' = $collectionId;
                                 'collectionName' = $collectionName;
                                 'siteId' = $siteId;
                                 'siteName' = $siteName;
                                 'description' = $description;
                                 'templateDeviceId' = $templateDeviceId;
                                 'templateDeviceName' = $templateDeviceName;
                                 'lastAutoAddDeviceNumber' = $lastAutoAddDeviceNumber;
                                 'enabled' = $enabled;
                                 'deviceCount' = $deviceCount;
                                 'deviceWithPVDCount' = $deviceWithPVDCount;
                                 'activeDeviceCount' = $activeDeviceCount; 
                                 'makActivateNeededCount' = $makActivateNeededCount;
                                 'autoAddPrefix' = $autoAddPrefix;
                                 'autoAddSuffix' = $autoAddSuffix;
                                 'autoAddZeroFill' = $autoAddZeroFill;
                                 'autoAddNumberLength' = $autoAddNumberLength;
                                 'role' = $role                                                                                                                     
                     }#result

                     $obj = New-Object -TypeName PSObject -Property $result
		     $obj.psobject.typenames.insert(0,'MorNando.PVS.GetPVScollection')
		     Write-Output $obj
        
                     $collection = $collection | select -skip 19
                     $c++

                } until($c -ge $counting)

}#process block finished

 END{
        try{
            remove-pssnapin *mcli* -ErrorAction 'silentlycontinue'
        }
        catch{}

 }#endblockfinished

}#end of Get-mPVScollection
