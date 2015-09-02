

$result = 0; #default Yes
$Name = Read-Host "Ime i prezime novog korisnika "
if (get-QADUser $Name) 
{
$title = "Detektovano isto ime!"
$message = "Korisnik sa istim imenom postoji - "+$Name+". Nastavi?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Nastavi"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Odustani"	
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
}

if (!($result))
{

$Filijala = Read-host @"
Ogranak Filijala:
 1. Becej
 2. BackaPalanka
 4. Beograd
 5. Jagodina (filijala 14)
 6. Kikinda
 7. Kragujevac (filijala 9)
 8. Kula 
 9. Nis (filijala 15)
10. NoviPazar (filijala 11)
11. Novisad
12. Pancevo
13. Ruma
14. Senta
15. Sid
16. Smederevo (filijala 12)
17. Sombor
18. SremskaMitrovica
19. StaraPazova
20. Subotica 
21. Uzice (filijala 10)
22. Valjevo (filijala 8)
23. Vranje (filijala 16)
24. Vrbas (filijala 7)
25. Vrsac 
26. Zrenjanin
 -	Centrala
 -	KorisnickiCentar
 -	Test
 -  Informatika
 -	Spoljni
(Unesi broj trazene filijale ili ime organizacione jedinice)
"@
$Password = "Novisad1" #inicijalni password
$Sec = $null
$Dis = $null


#$noemail = $null #kreiraj email

switch ( $Filijala )
{
	1
	{
	$OU = "OU=Korisnici,OU=Becej,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Becej,OU=Becej,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Becej,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	}
	2
	{
	$OU = "OU=Korisnici,OU=BackaPalanka,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=BackaPalanka,OU=BackaPalanka,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Backa Palanka,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	}
	4
	{
	$OU = "OU=Korisnici,OU=Beograd,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Beograd,OU=Beograd,OU=DDOR,DC=ddor,DC=local"
	}
	5
	{
	$OU = "OU=Korisnici,OU=Jagodina,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Jagodina,OU=Jagodina,OU=DDOR,DC=ddor,DC=local"
	}
	6 
	{
	$OU = "OU=Korisnici,OU=Kikinda,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Kikinda,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Kikinda,OU=Kikinda,OU=DDOR,DC=ddor,DC=local"
	}
	7
	{
	$OU = "OU=Korisnici,OU=Kragujevac,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Kragujevac,OU=Kragujevac,OU=DDOR,DC=ddor,DC=local"
	}
	8
	{
	$OU = "OU=Korisnici,OU=Kula,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Kula,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Kula,OU=Kula,OU=DDOR,DC=ddor,DC=local"
	}
	9
	{
	$OU = "OU=Korisnici,OU=Nis,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Nis,OU=Nis,OU=DDOR,DC=ddor,DC=local"
	}
	10
	{
	$OU = "OU=Korisnici,OU=NoviPazar,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=NoviPazar,OU=NoviPazar,OU=DDOR,DC=ddor,DC=local"
	}
	11
	{
	$OU = "OU=Korisnici,OU=NoviSad,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Novi Sad-1,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=NoviSad,OU=NoviSad,OU=DDOR,DC=ddor,DC=local"
	}
	12
	{
	$OU = "OU=Korisnici,OU=Pancevo,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Pancevo,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Pancevo,OU=Pancevo,OU=DDOR,DC=ddor,DC=local"
	}
	13
	{
	$OU = "OU=Korisnici,OU=Ruma,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Ruma,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Ruma,OU=Ruma,OU=DDOR,DC=ddor,DC=local"
	}
	14
	{
	$OU = "OU=Korisnici,OU=Senta,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Senta,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Senta,OU=Senta,OU=DDOR,DC=ddor,DC=local"
	}
	15
	{
	$OU = "OU=Korisnici,OU=Sid,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Sid,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Sid,OU=Sid,OU=DDOR,DC=ddor,DC=local"
	}
	16
	{
	$OU = "OU=Korisnici,OU=Smederevo,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Smederevo,OU=Smederevo,OU=DDOR,DC=ddor,DC=local"
	}
	17
	{
	$OU = "OU=Korisnici,OU=Sombor,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Sombor,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Sombor,OU=Sombor,OU=DDOR,DC=ddor,DC=local"
	}

	18
	{
	$OU = "OU=Korisnici,OU=SremskaMitrovica,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Srem,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=SremskaMitrovica,OU=SremskaMitrovica,OU=DDOR,DC=ddor,DC=local"
	}
	19
	{
	$OU = "OU=Korisnici,OU=StaraPazova,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Stara Pazova,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=StaraPazova,OU=StaraPazova,OU=DDOR,DC=ddor,DC=local"
	}
	20
	{
	$OU = "OU=Korisnici,OU=Subotica,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Subotica,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Subotica,OU=Subotica,OU=DDOR,DC=ddor,DC=local"
	}

	21
	{
	$OU = "OU=Korisnici,OU=Uzice,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Uzice,OU=Uzice,OU=DDOR,DC=ddor,DC=local"
	}

	22
	{
	$OU = "OU=Korisnici,OU=Valjevo,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Valjevo,OU=Valjevo,OU=DDOR,DC=ddor,DC=local"
	}
	23
	{
	$OU = "OU=Korisnici,OU=Vranje,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Vranje,OU=Vranje,OU=DDOR,DC=ddor,DC=local"
	}
	24
	{
	$OU = "OU=Korisnici,OU=Vrbas,OU=DDOR,DC=ddor,DC=local"
	$Dis = "CN=Vrbas,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Vrbas,OU=Vrbas,OU=DDOR,DC=ddor,DC=local"
	}
	25
	{
	$OU = "OU=Korisnici,OU=Vrsac,OU=DDOR,DC=ddor,DC=local" 
	$Dis = "CN=Vrsac,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
	$Sec = "CN=Vrsac,OU=Vrsac,OU=DDOR,DC=ddor,DC=local"
	}
	26
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
	Test
	{
	$OU = "OU=TESTSSO,OU=Nalozi i grupe koje kontrolisu administratori,OU=_Administrativni nalozi i grupe,OU=DDOR Administracija,DC=ddor,DC=local"
	}
	Informatika
	{
	$OU = "OU=Korisnici,OU=Informatika,OU=DDOR,dc=ddor,dc=local"
	}
	Spoljni
	{
	$OU = "OU=_SpoljniSaradnici,OU=DDOR Administracija,dc=ddor,dc=local"
	}

}

$msexc_db = Read-Host @"
Exchange database za korisnicki mailbox?
- POP3 (default)
- Headquarters
- ITS
- no
(mailboxdatabase)
"@
#####################################
$kurs = $Host.UI.RawUI.CursorPosition
#$a = $kurs.X 
$a = 0;
$b = $kurs.Y-1; #vrati se jedno mesto iznad
$koordinate = new-object System.Management.Automation.Host.Coordinates $a,$b
$Host.UI.RawUI.CursorPosition = $koordinate
##################################### vrati se na pocetak reda i jedan gore iznad

if ($msexc_db -eq "") {$msexc_db = "pop3"}
if ($msexc_db -eq "no") {Write-Host -nonewline "(mailboxdatabase): ";Write-Host -ForegroundColor Red "NO_EMAIL"}
else {Write-Host -nonewline "(mailboxdatabase): ";Write-Host -ForegroundColor Green $msexc_db}

#exit	#*** for testing

# Creating sAMAccountname
$n = $Name.Split()
$FirstName = $n[0]
$LastName = "$($n[1..$n.length])"
$Filijala = $n[2]

$Name = $FirstName + " "+$LastName
Write-Host "Kreiranje korisnickog naloga za: "$Name

$i = 0
[string]$fL = $FirstName[0..$i]
$fLs = $fL.Replace(" ","").ToLower()
$sAMAccountName = ($fLs + $LastName).ToLower()
$chk = Get-QADUser -SamAccountName $sAMAccountName
$SAM = $sAMAccountName

while ($chk -ne $null) 
{
	Write-Host "Korisnicko ime "$SAM" postoji!"
	$i++
	[string]$fL = $FirstName[0..$i]
	$fLs = $fL.Replace(" ","").ToLower()
	$sAMAccountName = ($fLs + $LastName).ToLower()
	$chk = Get-QADUser -SamAccountName $sAMAccountName
	$SAM = $sAMAccountName
} 

Write-Host -nonewline "Sledece raspolozivo korisnicko ime je: ";Write-Host -ForegroundColor Green $SAM

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
$user.SetInfo() 
Write-host "SetInfo1 - Kreiran Korisnik!"

$user.psbase.Invoke("SetPassword","Novisad1")
$user.pwdLastset=0
$user.psbase.InvokeSet("AccountDisabled",$false)
$user.put("UserAccountControl","512")
$user.SetInfo()
Write-Host "SetInfo2 - Setovan password i omogucen nalog!" 

#######################################
$xxx = @"
kreiranje mailboxa --- Peskoceno

if (!($noemail)) #ako je noemail setovan na 1,NE KREIRAJ EMAIL, preskoci sve ovo...
{

if ($MailServer -eq "POP3")
{

#check if session exist
#$chksess = Get-PSSession  | ? {($_.computername -like  "*e2010*") -and ($_.state -eq "Opened")}
#if ($chksess)  #if session exist and its open, user that session
#	{
#	$sessname = $chksess[0].Name #use first opened session and name it
#	} 
#else 
#	{

	#if no, create new session
	$uri = 'http://e2010-1.ddor.local/PowerShell'                                                                                                                         
	$sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos                                                           
#	}

[string]$sessuser = $sam
#$icmsess = icm $sess {@(get-user $sessuser)}
#	$icmsess.Name
Remove-PSSession $sess


}
else {
$user.mailNickname = $SAM
$user.msExchHomeServerName = $msExchHomeServerName_v
$user.homeMDB = $homeMDB_v
$user.mDBUseDefaults = $True
#$user.mDBOverHardQuotaLimit =  310000
#$user.mDBOverQuotaLimit = 300000
#$user.mDBStorageQuota = 290000
$user.setinfo()
Write-Host "SetInfo3 - Kreirani Exchange atributi!"

$ldap = "LDAP://"+$user.distinguishedName

$i = 1
while ($i -lt 1) # slicno kao sleep; ako je $i=1 nista se ne desava
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

	if($m % 2)
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
$a = 0
$koordinate = new-object System.Management.Automation.Host.Coordinates $a,$b
$Host.UI.RawUI.CursorPosition = $koordinate
Write-Host "                         "
$Host.UI.RawUI.CursorPosition = $koordinate
Write-Host "Kreiran je korisnicki mailbox:"$userAD.mail
Write-Host `n
$kreiran_email = $userAD.mail
}
}
#END kreiranje mailboxa ###################################################


else {
	$ldap = "LDAP://"+$user.distinguishedName
	$userAD = [ADSI]$ldap
	$kreiran_email = "NO_EMAIL"
	}
"@
#######################################


$ldap = "LDAP://"+$user.distinguishedName
$userAD = [ADSI]$ldap


#Ubacivanje korisnika u security grupu
if ($Sec -ne $null) {
Write-host "Ubacivanje korisnika u sec. grupe:"$Sec
$sgroupLDAP = "LDAP://"+$Sec
$ADsGroup = [ADSI]$sgroupLDAP
$adduserdn = $userAD.distinguishedName
$ADsGroup.member.add("$adduserdn")
$ADsGroup.Setinfo()
#Write-Host $sgroupLDAP
#Write-Host $adduserdn
#Write-Host $ADsGroup
Write-Host "SetInfo3 - Korisnik ubacen u sec.grupu"$ADsGroup.distinguishedName
}
#Ubacivanje korisnika u distributivnu grupu
if (($Dis -ne $null) -and ($noemail -ne "1"))
{
Write-host "Ubacivanje korisnika u dist. grupu:"$Dis
$dgroupLDAP = "LDAP://"+$Dis
$ADdGroup = [ADSI]$dgroupLDAP
$adduserdn = $userAD.distinguishedName
$ADdGroup.member.add("$adduserdn")
$ADdGroup.Setinfo()
Write-Host $dgroupLDAP
Write-Host $adduserdn
Write-Host $ADdGroup
Write-Host "SetInfo4 - Korisnik ubacen u grupu"$ADsGroup.distinguishedName
}

if ($msexc_db -ne "no")
{
while (!(get-user $SAM -ErrorAction "silentlycontinue"))
{
Write-host "." -NoNewline
Sleep 1
}
get-user $SAM | enable-mailbox -database $msexc_db
$mailbox = (get-mailbox $sam)
}

#mail administratoru
$emailFrom = "KreiranjeNaloga@DDOR.co.rs"
$emailTo = "milan.orlovic@ddor.co.rs"
$subject = "Kreiran novi korisnicki nalog - "+$userAD.samAccountName
$body = @"
Kreiran je novi korisnicki nalog: `n
$($userAD.distinguishedName) `n
Display Name:$($userAD.DisplayName) `n
Logon name:$($userad.sAMAccountName)`n
eMail:$($mailbox.PrimarySmtpAddress)`n
Korisnik ubacen u grupu:$($ADsGroup.Name)
"@
$smtpServer = "192.168.206.63"
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($emailFrom, $emailTo, $subject, $body)



if ($msexc_db -ne "no") #ako se kreira email, posalji korisniku....
{
#mail korisniku
#$emailFrom2 = "KreiranjeNaloga@DDOR"
#$emailTo2 = $userAD.mail
#$subject2 = "Kreiran korisnicki nalog -"+$userAD.samAccountName
$body2 = @"
Postovani/a,
Kreiran Vam je korisnicki nalog: `n
Display Name: $($userAD.DisplayName) `n
User Name: $($userad.sAMAccountName) (domenski nalog kojim se logujete na racunar)`n
eMail: $($userAD.mail)`n
"@
#$smtpServer = "192.168.206.63"
#$smtp2 = new-object Net.Mail.SmtpClient($smtpServer)
#$smtp2.Send($emailFrom2, $emailTo2, $subject2, $body2)
}


} #if (!($result))

break;
