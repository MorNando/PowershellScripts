###GLOBAL VARIABLES START###

$erroractionpreference = 'continue'

$global:scriptpath = split-path -parent $myinvocation.MyCommand.Definition
$global:WorkerWildCard = ""
$global:LongtermExcludedWorkers = get-content "$scriptpath\Config\LongTerm_ExcludedServers.txt"
$global:ShortTermExcludedWorkers = get-content "$scriptpath\Config\ShortTerm_ExcludedServers.txt"
$global:ProcessCount= "$scriptpath\Results\ProcessCount.csv"

###GLOBAL VARIABLES END###

Function get-ADworkerservers {

	import-module activedirectory
	$servers = get-adcomputer -filter { Name -like "*$WorkerWildCard*"} | select -expand name

	if ($LongTermExcludedWorkers -ne $null){

$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $LongtermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	}#longtermworkers ne null

	if ($ShortTermExcludedWorkers -ne $null){
$servers = Compare-Object -ReferenceObject $servers -DifferenceObject $ShorttermExcludedWorkers | Where-Object { $_.sideindicator -eq "<=" } | select -ExpandProperty inputobject

	} #SHORTTERMWORKERS NE NULL

	$servers

}#get-ADWorkerServers function end


Function Count-Processes {

Param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	$Servers

)
	$job = invoke-command -ComputerName $servers -ScriptBlock {
		$processname = $args[0]
		Try {
			$count = @(get-process -ea silentlycontinue "$processname").count
			new-object -typename psobject -property @{

				Server = $ENV:computername;
				RunningProcesses = $count
			}			
			
 		}#TRY MAIL PORT

		Catch {
	
		}#CATCH END

	} -asjob -throttlelimit 50 -erroraction continue -arg $processname ##invokecommandend


wait-job $job | out-null
#try{
	receive-job $job
#}#try end
#catch{

#}#catch end

}# Count-Processes Function End

###FUNCTIONS END###
$global:processname = read-host "What is the process you want to count?(without file extension)"
$servers = get-ADworkerservers
count-processes -servers $servers | select server,runningprocesses | export-csv $processcount -notypeinformation