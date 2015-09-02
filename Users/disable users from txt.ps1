
#$users = Get-Content D:\Scripts\_input\mart-ukidanje.txt
#$users = Get-Content 'D:\Temp\Ukidanje naloga 25Sep2012.txt'
#$users = Get-Content 'D:\Temp\Ukidanje 27Sep2012.txt'
#$users = Get-Content D:\Scripts\_input\Users\Brisanje11Dec2012.txt

#$users = Get-Content D:\Scripts\_input\Users\AccountExpired__need_disable.txt

#pretraga disable naloga
$users = get-qaduser -SizeLimit 0 |  ? {$_.Accountisexpired}


$report = "D:\Scripts\_Reports\ukidanje-report.txt"
$date = (get-date).AddDays(30).ToLongDateString()
$datum = (get-date).ToLongDateString()
#$datum_log = [string](get-date) -replace "/","-" -replace " ","_" -replace ":","-"
$datum_log = $datum -replace ", ","_" -replace " ","-" 
$report_name = "Ukidanje_naloga_"+$datum_log

#New-Item -ItemType File -Path "D:\Scripts\_Reports\" -Name $report_name | Set-Variable report

#$report = "D:\Scripts\_Reports\ukidanje-report"+$datum_log+".txt"

$i = 1
	foreach ($user in $users)
	{
	#_________________Kad je dato ime i prezime_____________za CSV
	#$fName = $user.ime
	#$lName = $user.prezime
	#$Ou = $user.orgou
		
	#$fullName = $fName+" "+$lName
	#$dName = $lName+" "+$fName
	#_______________________________________________________
	
	
	
	#$userAD = get-qaduser -DisplayName $dName
	#$userAD = get-qaduser -SamAccountName $user
	
	$userAD = get-qaduser $user
		$SAM = $userAD.samaccountname
		$dADName = $userAD.DisplayName
		$OUAD = $userAD.ParentContainer
		$desAD = $userAD.Description
	if ($userAD -ne $null)
	{
		if ($userAD.AccountIsDisabled) 
		{
		Write-Host -nonewline $i"."$user" postoji na domenu - ";Write-Host -nonewline -foreground Red $SAM ;Write-Host " "$dADName" - Ukinut nalog ranije!" 
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
		
			# Akcija 3
			Remove-QADGroupMember -Identity $groupAD -Member $userAD | Out-Null
		}
		
		$notes = "___old Notes: "+$userAD.Notes+"; ___new Notes: "+$datum+" ___Pripadnost grupama:"+$note+"; " 
		# Akcija 4
		$userAD | Set-QADUser -Notes $notes | out-null			
		
		# Akcija 5
		Disable-QADUser -Identity $SAM | Set-Variable disabled
		
		# Akcija 6
		Move-QADObject -Identity $SAM -NewParentContainer "OU=_disabled,OU=DDOR,DC=ddor,DC=local" | set-variable disabled

		Write-Host "Nalog" $userAD "je onemogucen, uklonjen iz GAL i prebacen u" $disabled.ParentContainerDN 
				Write-Host `n				
		$disabledou = $disabled.ParentContainerDN
		
		#_____ mail administratoru
	$emailFrom = "UkidanjeNaloga@DDOR.co.rs"
	$emailTo = "milan.orlovic@ddor.co.rs"
	$subject = "Ukinut korisnicki nalog - "+$userAD
$body = @"
Nalog $userAD je onemogucen,
uklonjen iz OAB i svih grupa (log o grupama zapisan u Notes),
prebacen u $disabledou
"@
	
	
	$smtpServer = "192.168.206.63"
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$smtp.Send($emailFrom, $emailTo, $subject, $body)

		
						
		}
	}
	else {
		 Write-Host $i"."$user "NE postoji na domenu!"
		 #Write-Host `n
		 }
	$i++	 
}


