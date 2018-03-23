<#
Author: Chris Morland
Date: 28/02/2018
Summary: This reboots a list of Server Names Using Runspaces
Version: 1.0

ChangeLog:
None
#>

#GlobalVariables

$Global:ScriptPath = split-path -parent $myinvocation.MyCommand.Definition
$Global:ServerNames = Get-Content $ScriptPath\Config\Servers_To_Reboot.txt
$Global:Results = "$ScriptPath\Results.txt"

Function Start-Runspace ( $CodeContainer, $ComputerName ){
    $runspacepool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1,10)
    $runspacepool.ApartmentState = "MTA"
    $runspacepool.Open()

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $threads = @()
    foreach ($c in $Computername) {

        $runspaceobject = [pscustomobject]@{
        
            Runspace = [powershell]::Create()
            Invoker = $null

        }
        $runspaceobject.Runspace.RunspacePool = $runspacepool
        $runspaceobject.Runspace.AddScript($CodeContainer) | Out-Null
        $runspaceobject.Runspace.AddArgument($c) | Out-Null
        $runspaceobject.Invoker = $runspaceObject.RunSpace.BeginInvoke()
        $threads += $runspaceobject
        $elapsed = $stopwatch.Elapsed
        Write-Host "Rebooting $c" -ForegroundColor "Yellow"

    }

    while ($threads.Invoker.IsCompleted -contains $false){}

    $elapsed = $stopwatch.Elapsed
    
    $threadresults = @()
    foreach ($t in $threads){

        $threadresults += $t.Runspace.EndInvoke($t.Invoker)
        $t.Runspace.Dispose()
    }

    $runspacepool.Close()
    $runspacepool.Dispose()

    return $threadresults

}

$RebootServerCodeContainer = {

    Param(
    [Parameter()]
    $ComputerName
    )
    
    $ComputerName = $ComputerName.toupper()
    $DateTime = Get-Date -Format "dd-MM-yy hh:mm"

    Try{
        $cmd = "\\$ComputerName\root\cimv2:Win32_OperatingSystem"
        $Server = [WMICLASS]$cmd

        if ($Server -ne $null){
            $ServerInstance = $Server.CreateInstance()
            $ReturnCode = $ServerInstance.reboot() | select -expand ReturnValue
            if ($ReturnCode -eq 0){
                $Result = "$ComputerName : Reboot successful at $DateTime by $Env:Username"
                return $Result
            }
            else{
                $Result = "$ComputerName : Reboot unsuccessful at $DateTime by $Env:Username"
                return $Result
            }
        }
    }
    Catch{
        if ($Server = $_.Exception.Message ){
            $Result = "$ComputerName : Reboot unsuccessful at $DateTime by $Env:Username"
			return $Result			
		}
    }
}
Start-Transcript -Path $Global:ScriptPath\Config\LastRunLog.log | Out-Null
Start-Runspace -CodeContainer $RebootServerCodeContainer -ComputerName $Global:ServerNames | Out-File -FilePath $Global:Results

Write-Host "`nReboots Complete! Check the results at the path:" -ForegroundColor "Green"
Write-Host "$Global:Results"

$HOST.UI.RawUI.Flushinputbuffer()
$HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
Stop-Transcript | Out-Null