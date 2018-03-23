Write-Host "Get Last Reboot Time"  -fore "yellow"
$OutputFile = ".\Results.csv"
$Computers = Get-Content -Path ".\Config\Devices.txt"
$count = $Computers.Count
$i = 0

if ((test-path "$OutputFile")){

    Get-ChildItem $OutputFile | Remove-Item
}

foreach ($Computer in $Computers){
    Write-host "Getting info for $computer : $i of $count"

    do {
        Get-Job -State Completed -HasMoreData $True | Receive-Job | select ComputerName, LastBootUpTime | Export-Csv -Path $OutputFile -NoTypeInformation -Append

    } While ( ( (Get-Job -State Running).Count ) -ge 30)

    Start-Job -Name $Computer -ScriptBlock {
        Param(
            [Parameter(Position=0)]
            $Computer
        )

        if ([system.io.directory]::Exists("\\$Computer\c$")){


                Get-WmiObject -ComputerName $Computer -Query "SELECT CSName,LastBootUpTime FROM Win32_OperatingSystem" | select @{n='ComputerName';e={$_.csname }}, @{n='LastBootUpTime';e={$_.ConvertToDateTime($_.LastBootupTime)}}

         }
         else{

                New-Object -TypeName PSObject -Property @{

                'ComputerName' = $Computer;
                'LastBootUpTime' = "Offline"

                }
          }
    } -ArgumentList $Computer | Out-Null
      

    $i++
}