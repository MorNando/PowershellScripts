###GLOBAL VARIABLES START###

$global:scriptpath = split-path -parent $myinvocation.MyCommand.Definition
$global:LongtermExcludedWorkers = get-content "$scriptpath\Config\LongTerm_ExcludedServers.txt"
$global:ShortTermExcludedWorkers = get-content "$scriptpath\Config\ShortTerm_ExcludedServers.txt"
$global:output = "$scriptpath\Results\FaultyPortServers.txt"
$global:errors = "$scriptpath\Results\uncontactable.txt"
$global:WorkerWildCard = ""
[int]$global:Port = read-host "Which port would you like to check?"
$global:Destination = read-host "What is the destination server ip address?"

###GLOBAL VARIABLES END###

###FUNCTIONS START###

Function Clean-OutputFile {

if (test-path $output){

remove-item $output

}#if end

}#Clean-OutputFile Function End

Function get-ADworkerservers {

	import-module activedirectory
	$servers = get-adcomputer -filter { Name -like "*$workerwildcard*"} | select -expand name

	if ($LongTermExcludedWorkers -ne $null){

$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $LongtermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	}#longtermworkers ne null

	if ($ShortTermExcludedWorkers -ne $null){
$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $ShorttermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	} #SHORTTERMWORKERS NE NULL

	$servers

}#get-ADWorkerServers function end


Function Test-TCPport {

Param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	$Servers

)
	$job = invoke-command -ComputerName $servers -ScriptBlock {
		$port=$Args[0]
		$destination=$Args[1]
		Try {
			$silent = New-Object Net.Sockets.TcpClient $destination, $port
			
 		}#TRY MAIL PORT

		Catch {
	
			if ($silent = $_.Exception.Message ){
				$Env:COMPUTERNAME
			}#IF SILENT
		}#CATCH END

	} -asjob -throttlelimit 50 -erroraction continue -arg $port,$destination ##invokecommandend


wait-job $job | out-null
#try{
	receive-job $job
#}#try end
#catch{

#if ($job = $_.Exception.Message ){
			#	$_.Exception.Source | out-file $errors -append -confirm:$false
			#}#IF SILENT


#}#catch end

}# Test-TCPport Function End

###FUNCTIONS END###

###CODE START###

Clean-OutputFile
$Servers = Get-ADWorkerServers 
Test-TCPport -Servers $Servers | out-file $output -append -confirm:$false

###CODE END###