$global:scriptpath = split-path -parent $myinvocation.MyCommand.Definition
$global:masterimagepath = get-content "$scriptpath\Config\Master_Image_Path.txt"
$global:pvsserver = $Env:COMPUTERNAME
$global:pvsfolder = ($masterimagepath).Split("\") | select -Last 1
$global:robocopylog = "$scriptpath\Logs\Robocopy"

get-childitem C:\Scripts\Robocopy-Image\Logs | Where-Object {$_.lastwritetime -lt (get-date).adddays(-1)} | remove-item 

Function Begin-Robocopy { 
    Param(

        $image
    )
	if ($pvsserver -notlike "*PVS*"){
		
		Write-host "Please run on each PVS server individually" -fore red
	
	}else{

            if ($pvsserver -like "*Farm1*"){
                $pvsfolder = ($pvsfolder).replace("PVSXX","PVSYY")
                $pvsdestination = "\\$pvsserver\$pvsfolder"
            }else{

		    $pvsfolder = ($pvsfolder).replace("PVSYY","PVSXX")
		    $pvsdestination = "\\$pvsserver\$pvsfolder"

	        }
	    
            $pvpvalid = Test-Path "$pvsdestination\$image.pvp"
            $vhdvalid = Test-Path "$pvsdestination\$image.vhd"

            if ($pvpvalid){

                $timestamppvp = (Get-Item "$pvsdestination\$image.pvp" | select -expand LastWriteTime).year

                if($timestamppvp -lt "2012"){$pvpvalid = $false}

            }#timestamppvp file corrupt check

            if ($vhdvalid){
            
                $timestampvhd = (Get-Item "$pvsdestination\$image.vhd" | select -expand LastWriteTime).year

                if($timestampvhd -lt "2012"){$vhdvalid = $false}

            }#timestampvhd file corrupt check

	    [FLOAT]$freespace = "{0:N2}" -f ((get-psdrive | Where-Object {$_.name -eq "V"} | select -expand Free) / 1gb)
	    $pvp = get-childitem $masterimagepath\$image.pvp | select -expand length
	    $vhd = get-childitem $masterimagepath\$image.vhd | select -expand length
	    [FLOAT]$total = "{0:N2}" -f (($pvp + $vhd) / 1GB)
	    $total = $total + 10

	    if ($freespace -le $total){

	    	Write-Host "`nNot enough disk space on the drive" -Fore Red
            	$HOST.UI.RawUI.Flushinputbuffer()
    		$HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
    		Break

	    }else{Write-Host "`nCheck:Diskspace is fine`n" -fore Green}
	    
            if (($pvpvalid -eq $false) -or ($vhdvalid -eq $false)){

                    robocopy "$masterimagepath" "$pvsdestination" "$image.pvp" /COPY:DATSOU /R:1 /W:30 | Out-File "$robocopylog`_$pvsserver_$image.log"
                    robocopy "$masterimagepath" "$pvsdestination" "$image.vhd" /COPY:DATSOU /R:1 /W:30  | Out-File "$robocopylog`_$pvsserver_$image.log" -append

            
            }#if pvpvalid or vhdvalid are false robocopy begins

}
}#Begin-Robocopy function end

Function import-PVSimage {
Invoke-Command -ComputerName Server01 {
    $image = $args
    Import-Module mornandoPVStools


} -ArgumentList $image

}# import-PVSimage function end


    do{
        Clear-Host
        $i = 1
        Write-Host "Your master image path is $masterimagepath" -ForegroundColor Green
        Write-Host '(This can be edited in the Config folder)'`n
        $imagelist = Get-ChildItem "$masterimagepath" | Where-Object {$_.name -like "*XA*" } | select @{ n="imagename"; e={$_.name.split(".") | select -First 1 }} -Unique | ForEach-Object {

        New-Object -TypeName PSObject -Property @{

                'Index' = $i
                'ImageName' = ($_ | select -ExpandProperty ImageName )

            } #newobject

        $i++
        } 

        $imagelist | select index,ImageName | ft -AutoSize

        try{

        	$imagecount = $imagelist.count
        	[int]$ImageNumber = read-host "Enter the index number of your new master image "

		if ($imageNumber -lt 1 -or $imageNumber -gt $imagecount){
       	    		"The image number is invalid. Press any key to select again."
    			Break
		}

	$Image = ($Imagelist[($ImageNumber - 1)]).ImageName
        if ($image -notlike "*XA65*"){
            "The image number is invalid. Press any key to select again."
            Break
        }#if image not like xa65

        }Catch{
                
            "The image number is invalid. Press any key to select again."
            Break
        
        }#catch end


        $question = "You are provisioning {0}. Is this correct? y/n" -f $image
        $ans = Read-Host $question
        
    }until($ans -eq "yes" -or $ans -eq "y")#do until end

write-host "Copying $image to pvsserver $pvsserver if required. Check the logs folder for progress." -ForegroundColor Yellow
Begin-Robocopy -image $image