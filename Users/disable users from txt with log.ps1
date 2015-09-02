

###
# 	NEDOVRSENO!
###


#$users = Get-Content D:\Scripts\_input\mart-ukidanje.txt
#$users = Import-Csv "D:\Scripts\_input\Users\disable-27Sep2011.txt" -Delimiter ";"
#$users = Get-Content D:\Scripts\_input\Users\AccountExpired__need_disable_test.txt
$users = get-qaduser -SizeLimit 0 |  ? {$_.Accountisexpired}

$date = (get-date).AddDays(30).ToLongDateString()
$datum = (get-date).ToLongDateString()
#$datum_log = [string](get-date) -replace "/","-" -replace " ","_" -replace ":","-"
$datum_log = $datum -replace ", ","_" -replace " ","-" 
$report_name = "Ukidanje_naloga_"+$datum_log+".txt"
#$logname = "ukidanje_naloga_"+$datum_log+".txt"
New-Item -ItemType File -Path "D:\Scripts\_Reports\" -Name $logname | Set-Variable report

#$report = "D:\Scripts\_Reports\ukidanje-report"+$datum_log+".txt"

$i = 1
	foreach ($user in $users)
	{
	Write-Output $user >> $report
	#_________________Kad je dato ime i prezime_____________za CSV
	#$fName = $user.ime
	#$lName = $user.prezime
	#$Ou = $user.orgou
	
	#$fullName = $fName+" "+$lName
	#$dName = $lName+" "+$fName
	#_______________________________________________________
	
	$userAD = get-qaduser -SamAccountName $user
	#$userAD = get-qaduser -DisplayName $dName
		$SAM = $userAD.samaccountname
		$dADName = $userAD.DisplayName
		$OUAD = $userAD.ParentContainer
		$desAD = $userAD.Description
	if (($userAD -ne $null) -and ($ouad -like "*$ou*"))
	{
		if ($userAD.AccountIsDisabled) 
		{
		Write-Host -nonewline $i"."$user" postoji na domenu - ";Write-Host -nonewline -foreground Red $SAM ;Write-Host " "$dADName" - Ukinut nalog!" 
		$o1 = "Nalog "+$dADNAme+"je vec onemogucen!"
		Write-Output $o1 >> $report
		$o1 = $null
		}
		else 
		{
		Write-Host -nonewline $i"."$user" postoji na domenu - ";Write-Host -nonewline -foreground green $SAM ;Write-Host " "$dADName,$OU 
		
		# Akcija 1
		Set-QADUser -Identity $SAM -Description "delete after $date; former OU: $($OU); previous description: $($desc)" | Out-Null 
				
		# Akcija 2
		Set-QADObject -Identity $SAM -ObjectAttributes @{msExchHideFromAddressLists=$true}	| Out-Null
		
		$notes=$null
		$note=$null
		foreach ($group in $userAD.MemberOf)
		{
			$groupAD = get-qadgroup $group
			$gName = $groupAD.Name 
			$gType = $groupAD.GroupType 
		
			$note += $gName+" - "+$gtype+", "
			Write-Host "Removing"$userad.Name"from "$groupAD
			$o2 = "Removing"+$userad.Name+"from "+$groupAD
			Write-Output $o2 >> $report
			
			# Akcija 3
			Remove-QADGroupMember -Identity $groupAD -Member $userAD | Out-Null
			$o2 = $null
		}
		
		$notes = "___old Notes: "+$userAD.Notes+"; ___new Notes: "+$datum+" ___Pripadnost grupama:"+$note+"; " 
		
		# Akcija 4
		$userAD | Set-QADUser -Notes $notes | out-null			
		#Write-Output $notes >> $report
		
		# Akcija 5
		Disable-QADUser -Identity $SAM | Set-Variable disabled
		$o3 = "Nalog "+$userAD.Name+" ("+$SAM+") je sada onemogucen!"
		Write-Output $o3 >> $report
		$o3 = $null		
		
		# Akcija 6
		Move-QADObject -Identity $SAM -NewParentContainer "OU=_disabled,OU=DDOR,DC=ddor,DC=local" | set-variable disabled

		Write-Host "Nalog" $userAD "je onemogucen, uklonjen iz GAL i prebacen u" $disabled.ParentContainerDN 
		Write-Host `n	
		
		#_____ mail administratoru
	$emailFrom = "UkidanjeNaloga@DDOR.co.rs"
	$emailTo = "milan.orlovic@ddor.co.rs"
	$subject = "Ukinut korisnicki nalog - "+$username
$body = @"
Nalog $username je onemogucen,
uklonjen iz OAB i svih grupa (log o grupama zapisan u Notes),
prebacen u $disabledou
"@
	
	$smtpServer = "192.168.206.15"
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$smtp.Send($emailFrom, $emailTo, $subject, $body)

		
						
		}
	}
	else {
		 Write-Host $i"."$user "NE postoji na domenu!"
		 #Write-Host `n
		 $o4="Ne postoji na domenu!"
		 Write-Output $o4 >> $report
		 $o4 = $null
		 }
	$i++	 
}