#
#Kreiranje korisnickih naloga i mailboxova iz csv fajla
#

$log = "D:\Scripts\_Reports\Kreiranje Naloga\Kreiranje-Naloga-Prodavci.txt"

$users = Import-Csv "D:\Scripts\_input\Kreiranje Prodavaca - test.csv"
$Password = "Novisad1" #inicijalni password
$existing = New-Object System.Collections.ArrayList
$allusers = Get-QADUser -SizeLimit 0
#msExchHomeServerName_v = "/O=DDOR Novi Sad ad/OU=LIMAN/cn=Configuration/cn=Servers/cn=HERMES"
#$homeMDB_v ="CN=Ostali,CN=First Storage Group,CN=InformationStore,CN=HERMES,CN=Servers,CN=LIMAN,CN=Administrative Groups,CN=DDOR Novi Sad ad,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=ddor,DC=local"

$msExchHomeServerName_v = "/O=DDOR Novi Sad ad/OU=LIMAN/cn=Configuration/cn=Servers/cn=EXC2003"
$homeMDB_v = "CN=Sales,CN=First Storage Group,CN=InformationStore,CN=EXC2003,CN=Servers,CN=LIMAN,CN=Administrative Groups,CN=DDOR Novi Sad ad,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=ddor,DC=local"

#$server = "192.168.206.15" 
#$client = new-object system.net.mail.smtpclient $server 

#$smtpServer = "192.168.206.15"
#$smtp = new-object Net.Mail.SmtpClient($smtpServer)

foreach ($user in $users)
{

# Creating sAMAccountname
$FirstName = ($user.Ime).Trim()
$LastName = ($user.Prezime).Trim()
$desc = $user.Opis
$OUcsv = $user.Filijala.Trim()
#$jmbg = $user.JMBG

switch ( $OUcsv )
{
	Becej
	{
		$OU = "OU=Korisnici,OU=Becej,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Becej,OU=Becej,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Becej,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	}
	BackaPalanka
	{
		$OU = "OU=Korisnici,OU=BackaPalanka,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=BackaPalanka,OU=BackaPalanka,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Backa Palanka,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	}
	BackaTopola
	{
		$OU = "OU=Korisnici,OU=BackaTopola,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=BackaTopola,OU=BackaTopola,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=BackaTopola,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	}
	Beograd
	{
		$OU = "OU=Korisnici,OU=Beograd,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Beograd,OU=Beograd,OU=DDOR,DC=ddor,DC=local"
	}
	Jagodina
	{
		$OU = "OU=Korisnici,OU=Jagodina,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Jagodina,OU=Jagodina,OU=DDOR,DC=ddor,DC=local"
	}
	Kikinda 
	{
		$OU = "OU=Korisnici,OU=Kikinda,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Kikinda,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Kikinda,OU=Kikinda,OU=DDOR,DC=ddor,DC=local"
	}
	Kragujevac
	{
		$OU = "OU=Korisnici,OU=Kragujevac,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Kragujevac,OU=Kragujevac,OU=DDOR,DC=ddor,DC=local"
	}
	Kula
	{
		$OU = "OU=Korisnici,OU=Kula,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Kula,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Kula,OU=Kula,OU=DDOR,DC=ddor,DC=local"
	}
	Nis
	{
		$OU = "OU=Korisnici,OU=Nis,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Nis,OU=Nis,OU=DDOR,DC=ddor,DC=local"
	}
	NoviPazar
	{
		$OU = "OU=Korisnici,OU=NoviPazar,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=NoviPazar,OU=NoviPazar,OU=DDOR,DC=ddor,DC=local"
	}
	NoviSad
	{
		$OU = "OU=Korisnici,OU=NoviSad,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Novi Sad-1,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=NoviSad,OU=NoviSad,OU=DDOR,DC=ddor,DC=local"
	}
	Pancevo
	{
		$OU = "OU=Korisnici,OU=Pancevo,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Pancevo,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Pancevo,OU=Pancevo,OU=DDOR,DC=ddor,DC=local"
	}
	Ruma
	{
		$OU = "OU=Korisnici,OU=Ruma,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Ruma,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Ruma,OU=Ruma,OU=DDOR,DC=ddor,DC=local"
	}
	Senta
	{
		$OU = "OU=Korisnici,OU=Senta,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Senta,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Senta,OU=Senta,OU=DDOR,DC=ddor,DC=local"
	}
	Sid
	{
		$OU = "OU=Korisnici,OU=Sid,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Sid,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Sid,OU=Sid,OU=DDOR,DC=ddor,DC=local"
	}
	Smederevo
	{
		$OU = "OU=Korisnici,OU=Smederevo,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Smederevo,OU=Smederevo,OU=DDOR,DC=ddor,DC=local"
	}
	Sombor
	{
		$OU = "OU=Korisnici,OU=Sombor,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Sombor,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Sombor,OU=Sombor,OU=DDOR,DC=ddor,DC=local"
	}

	SremskaMitrovica
	{
		$OU = "OU=Korisnici,OU=SremskaMitrovica,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Srem,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=SremskaMitrovica,OU=SremskaMitrovica,OU=DDOR,DC=ddor,DC=local"
	}
	StaraPazova
	{
		$OU = "OU=Korisnici,OU=StaraPazova,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Stara Pazova,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=StaraPazova,OU=StaraPazova,OU=DDOR,DC=ddor,DC=local"
	}
	Subotica
	{
		$OU = "OU=Korisnici,OU=Subotica,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Subotica,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Subotica,OU=Subotica,OU=DDOR,DC=ddor,DC=local"
	}

	Uzice
	{
		$OU = "OU=Korisnici,OU=Uzice,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Uzice,OU=Uzice,OU=DDOR,DC=ddor,DC=local"
	}

	Valjevo
	{
		$OU = "OU=Korisnici,OU=Valjevo,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Valjevo,OU=Valjevo,OU=DDOR,DC=ddor,DC=local"
	}


	Vranje
	{
		$OU = "OU=Korisnici,OU=Vranje,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Vranje,OU=Vranje,OU=DDOR,DC=ddor,DC=local"
	}
	Vrbas
	{
		$OU = "OU=Korisnici,OU=Vrbas,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Vrbas,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Vrbas,OU=Vrbas,OU=DDOR,DC=ddor,DC=local"
	}
	Vrsac
	{
		$OU = "OU=Korisnici,OU=Vrsac,OU=DDOR,DC=ddor,DC=local" 
		$Dis = "CN=Vrsac,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Vrsac,OU=Vrsac,OU=DDOR,DC=ddor,DC=local"
	}
	Zrenjanin
	{
		$OU = "OU=Korisnici,OU=Zrenjanin,OU=DDOR,DC=ddor,DC=local"
		$Dis = "CN=Zrenjanin,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
		$Sec = "CN=Zrenjanin,OU=Zrenjanin,OU=DDOR,DC=ddor,DC=local"
	}
	Centrala 
	{
		$OU = "OU=Korisnici,OU=Centrala,OU=DDOR,DC=ddor,DC=local"
	}
	KorisnickiCentar
	{
		$OU = "OU=Korisnici,OU=KorisnickiCentar,OU=DDOR,DC=ddor,DC=local"
		$Sec =  "CN=KorisnickiCentar,OU=KorisnickiCentar,OU=DDOR,DC=ddor,DC=local"
	}
}#end switch

$Name = $FirstName+" "+$LastName
$lName = $LastName+" "+$FirstName
$out1 = "Kreiranje korisnickog naloga za: " +$Name+ "," +$OUcsv+ " --> " +$OU
Write-Output $out1 >> $log
$out1=$null
$i = 0
[string]$fL = $FirstName[0..$i]
$fLs = $fL.Replace(" ","").ToLower()
$SAM = ($fLs + $LastName).ToLower()
$chk_sam = $allusers | where {$_.SamAccountName -eq $SAM}
$chk_name = $allusers | where {$_.DisplayName -eq $lName}


if ($chk_name -eq $null) #Korisnik po imenu ne postoji na domenu
{
 $out2="Korisnik sa imenom " + $lName + " ne postoji na domenu, kreiram nalog..." 
 Write-Output $out2 >> $log
 $out2=$null
 while ($chk_sam -ne $null) #proveri da li postoji zeljeno korisnicko ime, ako postoji...
 {
 	$out3= "Korisnicko ime " +$SAM+ " je vec zauzeto - " +(Get-QADUser $SAM).Name+ "," +(Get-QADUser $SAM).samaccountname+ "," +(Get-QADUser $SAM).parentcontainer
	Write-Output  $out3 >> $log
	$out3 = $null
	#Write-Host -NoNewline "	Korisnicko ime ";
	#Write-Host -NoNewline -ForegroundColor Red $SAM;
	#Write-Host " je vec zauzeto - "(Get-QADUser $SAM).Name","(Get-QADUser $SAM).samaccountname","(Get-QADUser $SAM).parentcontainer
	
	$i++
	[string]$fL = $FirstName[0..$i]
	$fLs = $fL.Replace(" ","").ToLower()
	$SAM = ($fLs + $LastName).ToLower()
	$chk_sam = $allusers | where {$_.SamAccountName -eq $SAM}
	
} 

$out4 = "Sledece raspolozivo korisnicko ime je: " +$SAM
Write-Output  $out4 >> $log
$out4 = $null
#Write-Host -nonewline "	Sledece raspolozivo korisnicko ime je: ";Write-Host -ForegroundColor Green $SAM

#kreiranje naloga za odabrani $SAM i $OU
$Container = [ADSI]"LDAP://$OU"
$user = $Container.Create("user","cn=$Name") 
$user.put("sn",$LastName)
$user.put("givenName",$FirstName)
$user.put("displayName",$Lastname + " "+$Firstname) 
$root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
$suffix = $root.defaultNamingContext -replace "dc=" -replace ",","."

$upn = "$SAM@$suffix"
$user.put("userPrincipalName",$upn)
$user.put("sAMAccountName",$SAM)
$user.put("description",$desc)
$user.SetInfo() 
Write-host "SetInfo1 - Kreiran Korisnik"

$user.psbase.Invoke("SetPassword","Novisad1")
$user.pwdLastset=0
$user.psbase.InvokeSet("AccountDisabled",$false)
$user.SetInfo()
Write-Host "SetInfo2 - Setovan password i omogucen" 


$user.mailNickname = $SAM
$user.msExchHomeServerName = $msExchHomeServerName_v
$user.homeMDB = $homeMDB_v
$user.mDBUseDefaults = $True
#$user.mDBOverHardQuotaLimit =  310000
#$user.mDBOverQuotaLimit = 300000
#$user.mDBStorageQuota = 290000

$user.setinfo()
Write-Host "SetInfo3 - Mailbox atributi"


$ldap = "LDAP://"+$user.distinguishedName

$i = 1
while ($i -lt 1)
{
	Write-Host -noNewLine "."
	Sleep 1
	$i++
}
Write-Host "`n"

$userAD = [ADSI]$ldap
Write-Host -noNewLine "Kreiranje mailboxa"
$emp = "        "
$kurs = $Host.UI.RawUI.CursorPosition
$a = $kurs.X
$b = $kurs.Y
$koordinate = new-object System.Management.Automation.Host.Coordinates $a,$b
$m = 1
while (!($userAD.mail)) 
{

	if($m%2)
	{
		$dot = "."
		$mm = 1
		while($mm -lt 6)
		{
			$Host.UI.RawUI.CursorPosition = $koordinate
			Write-Host -noNewLine $dot
			Sleep -milliseconds 500
			$dot = $dot + "."
			$mm++
		}
	}
	else 
	{
	$dot = $emp;
	$Host.UI.RawUI.CursorPosition = $koordinate
	Write-Host -noNewLine $dot 
	Sleep -milliseconds 500
	}
$userAD = [ADSI]$ldap
$m++
}#end while 
$a=0
$koordinate = new-object System.Management.Automation.Host.Coordinates $a,$b
$Host.UI.RawUI.CursorPosition = $koordinate
Write-Host "                         "
$Host.UI.RawUI.CursorPosition = $koordinate
Write-Host "Kreiran je korisnicki mailbox:"$userAD.mail
Write-Host `n

$smtpServer = "192.168.206.15"
$emailTo = "milan.orlovic@ddor.co.rs"
$emailFrom = "KreiranjeNaloga@DDOR.co.rs"

#email administratoru
$emailFrom = "KreiranjeNaloga@DDOR"
$emailTo = "milan.orlovic@ddor.co.rs"
$subject = "Kreiran novi korisnicki nalog - "+$userAD.samAccountName
$body = @"
Kreiran je novi korisnicki nalog: $($userAD.distinguishedName) `n
Display Name:$($userAD.DisplayName) `n
Logon name:$($userad.sAMAccountName)`n
eMail:$($userAD.mail)`n
"@
Send-MailMessage -SmtpServer $smtpServer -Body $body -To $emailTo -From $emailFrom -Subject $subject

Write-Output `n >> $log
Write-Output $body >> $log
Write-Output "_______________________________" >> $log

####check DC in site against created user
#$dclist = "zevs","novisad2","vm-posejdon"
#Write-Host "...checking "$userAD.distinguishedName
#foreach ($dc in $dclist)
#{
#Write-Host -NoNewline $dc
#[string]$objLDAP = "LDAP://$dc/"+$userAD.distinguishedName 
#Write-Host -NoNewline "   "$objLDAP
#$object = [adsi]$objLDAP
#if ($object.name) {Write-Host -foregroundcolor Green " has object,replicated!" } else {Write-Host -ForegroundColor Red " not replicated!"}
#}

$wi=15
while ($wi -ge 1) 
{
Write-Host -NoNewline `r$wi;
Sleep 1
Write-Host -NoNewline `r"  "
$wi--
}

#Write-Host `n
#email korisniku
$emailTo2 = $userAD.mail
$subject2 = "Kreiran korisnicki nalog - "+$userAD.samAccountName
$body2 = @"
Kreiran Vam je korisnicki nalog: $($userAD.distinguishedName) `n
Display Name:$($userAD.DisplayName) `n
Logon Name:$($userad.sAMAccountName)`n
eMail:$($userAD.mail)`n
"@
Send-MailMessage -SmtpServer $smtpServer -Body $body2 -To $emailTo2 -From $emailFrom -Subject $subject2 

}

else #korisnik vec postoji, predji na drugog i napravi spisak za dodatnu proveru!
{
$out5 = "Korisnik sa imenom " +$chk_name.DisplayName+ "|" +$chk_name.parentcontainer+  "postoji, upisujem u niz!"
Write-Output $out5 >> $log
Write-Output "_______________________________" >> $log
$out5 = $null
#Write-Host "Korisnik sa imenom "$chk_name.DisplayName "|" $chk_name.parentcontainer  "postoji, upisujem u niz!"
$null_p = $existing.Add($chk_name)
}

#Write-Host `n
Write-Output `n >> $log
}#end foreach

Write-Output `n >> $log
Write-Output "____________________" >> $log
Write-Output "Postojeci korisnici:" $existing >> $log
#Write-Host "Postojeci korisnici:"
$existing



