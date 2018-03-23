$Computernames = get-content ".\Config\ComputerNames.txt"
$Cred = Get-Credential
$output = ".\Results.txt"
$shortdomain = ""
$domainfqdn = ""
$localusername = ""
$localpassword = ""
$UserName = $Cred.UserName
$SecurePassword = $Cred.Password
$UnsecurePassword = (New-Object PSCredential "user",$SecurePassword).GetNetworkCredential().Password 

Write-Host "Rejoining Machine to $domain Domain"  -fore "green"

Foreach ($Computername in $Computernames){

	cmdkey /add:$ComputerName /user:$Computername\$localusername /pass:$localpassword
    $check = netdom verify $Computername
    $check = $check | Where-Object { $_ -like "*$shortdomain is invalid*"}
    $check2 = $check | Where-Object { $_ -like "*$shortdomain has been verified*"} 
    if ($check -like "*$shortdomain is invalid*"){

        c:\Install\Utilities\PSTools\psexec.exe \\$ComputerName netdom remove $Computername /domain:$domainfqdn /userd:$UserName /passwordd:$UnsecurePassword
        Write-host "sleeping"
        start-sleep 90 -Verbose
        Write-host "wake up"
	    c:\Install\Utilities\PSTools\psexec.exe \\$ComputerName netdom join $Computername /domain:$domainfqdn /userd:$UserName /passwordd:$UnsecurePassword /reboot
        Write-Host "Starting sleep while reboot takes place on $computername"
        Start-Sleep 240 -Verbose
        $verify = netdom verify $Computername
        $verifyfinal = $verify | Where-Object { $_ -like "*$shortdomain has been verified*" }

        if ($verifyfinal -like "*$shortdomain has been verified*"){
            $ComputerName + ": Successfully fixed" | Out-File -FilePath $output -Append
        }
        else{
            $verifyfinal = $verify | Where-Object { $_ -like "*$shortdomain is invalid*" }
            if ($verifyfinal -like "*$shortdomain is invalid*"){
                $ComputerName + ": Failed on domain join" | Out-File -FilePath $output -Append
            }
            else{
                 $ComputerName + ": Unknown" | Out-File -FilePath $output -Append
            }
        }

    }
    elseif($check2 -like "*$shortdomain has been verified*"){
            $ComputerName + ": Already Successful" | Out-File -FilePath $output -Append
    }    
    else{

        $ComputerName + ": Unknown" | Out-File -FilePath $output -Append

    }


}

Write-Host "COMPLETE"  -fore "green"
Read-host "Press any key to continue"
