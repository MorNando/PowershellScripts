$erroractionpreference = 'silentlycontinue'

###GLOBAL VARIABLES START###

$erroractionpreference = 'continue'
$global:WorkerWildCard = ""
$global:scriptpath = split-path -parent $myinvocation.MyCommand.Definition
$global:WorkerDiskInfo = "$scriptpath\Results\WorkerDiskInfo.csv"
$global:LongtermExcludedWorkers = get-content "$scriptpath\Config\LongTerm_ExcludedServers.txt"
$global:ShortTermExcludedWorkers = get-content "$scriptpath\Config\ShortTerm_ExcludedServers.txt"


###GLOBAL VARIABLES END###


###FUNCTIONS START###

Function get-ADworkerservers {

	import-module activedirectory
	$servers = get-adcomputer -filter { Name -like "*$global:WorkerWildCard*"} | select -expand name

	if ($LongTermExcludedWorkers -ne $null){

$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $LongtermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	}#longtermworkers ne null

	if ($ShortTermExcludedWorkers -ne $null){
$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $ShorttermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	} #SHORTTERMWORKERS NE NULL

	$servers

}#get-ADWorkerServers function End

Function Get-WorkerDiskInfo {
    
    get-wmiobject -ComputerName $workers -class win32_logicaldisk -Filter 'DriveType=3' | ForEach-Object {
    
    New-Object -TypeName psobject -Property @{
	    
	    'ServerName' = $_.__SERVER
            'DriveLetter' = $_.deviceid
	    'UsedSpace(GB)' = (("{0:N2}" -f ($_.size / 1gb )) - ("{0:N2}" -f ($_.freespace / 1gb)))
            'Size(GB)' = "{0:N2}" -f ($_.size / 1gb )
            'FreeSpace(GB)' = "{0:N2}" -f ($_.freespace / 1gb)
            '%Used' = "{0:N2}" -f (  (("{0:N2}" -f ($_.size / 1gb )) - ("{0:N2}" -f ($_.freespace / 1gb))) / ("{0:N2}" -f ($_.size / 1gb ) ) * 100 )

    } | where-object {$_.DriveLetter -eq 'C:' -or $_.DriveLetter -eq 'D:' } | select ServerName,DriveLetter,'UsedSpace(GB)','Size(GB)','Freespace(GB)',%Used

    } 
} # Get-WorkerDiskInfo Function End

$workers = get-ADworkerservers

Get-WorkerDiskInfo | export-csv $WorkerDiskInfo -NoTypeInformation