## upis opisa i propadnost grupama korisnika u atribut Notes
##uklanjanje disabled usera iz svih grupa

foreach ($user in (Get-QADUser -SearchRoot "ddor.local/ddor/_disabled" | Sort-Object Name))
#$datum = (get-date).ToLongDateString()
#foreach ($user in (Get-QADUser -samaccountname alukic))
{
	$userAD = Get-QADUser $user
	Write-host $userAD.Name " Disabled:"$userAD.AccountIsDIsabled
	#Write-host $userAD.Notes
	$note,$notes=$null
	foreach ($group in $userAD.MemberOf)
	{
	
		$groupAD = get-qadgroup $group
		#$gName = $groupAD.Name 
		#$gType = $groupAD.GroupType 
		#$group.GroupType
		#$note += $gName+" - "+$gtype+", "
		Write-Host "Removing "$userAD.Name" from "$groupAD.Name
		Remove-QADGroupMember -Identity $groupAD -Member $userAD -Confirm
		
	}
	
	#$prevdesc = $userAD.Description
	#$notes = "Stari Notes:"+$userAD.notes+"`n"+" "+$datum+$prevdesc+": Pripadnost grupama:"+$note+"; "
	#$userAD | Set-QADUser -Notes $notes
	
	#$notes 
	Write-Host "____________"
	#Write-host $userAD.Notes
	
}
