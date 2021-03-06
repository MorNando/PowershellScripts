###GLOBAL VARIABLES START###

$erroractionpreference = 'continue'

$global:scriptpath = split-path -parent $myinvocation.MyCommand.Definition
$global:Farm1errorlogpath = "$scriptpath\results\Farm1_error_logs.csv"
$global:Farm2errorlogpath = "$scriptpath\results\Farm2_error_logs.csv"
$global:WorkerWildCard = ""
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

}#get-ADWorkerServers function end

Function get-Farm1Events {
Param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	$Computer

)
		Get-eventlog -logname system -EntryType error -newest 60 -ComputerName $computer| select machinename,Eventid,message,timegenerated


}#Get-Farm1Events Function End

Function Get-Farm2Events {
Param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	$Computer

)
	Get-eventlog -logname system -EntryType error -newest 60 -ComputerName $computer| select machinename,Eventid,message,timegenerated

}#Get-Farm2Events Function End

###FUNCTIONS END###

###CODE START###

write-host "Gathering worker information from AD..." -fore "Yellow"
$servers = get-ADworkerservers

write-host "gathering events from Farm1 Workers..." -fore "Yellow"
$servers | where-object {$_ -like "Farm1*$global:WorkerWildCard*"} | foreach-object { Get-Farm1Events $_ } | export-csv $Farm1errorlogpath -notypeinformation

write-host "gathering events from Farm2 Workers..." -fore "Yellow"
$servers | where-object {$_ -like "Farm2*$global:WorkerWildCard*"} | foreach-object {Get-Farm1Events $_ } | export-csv $Farm2errorlogpath -notypeinformation


###CODE END###