
$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Columns.Item('A').ColumnWidth = 15
$c.Columns.Item('B').ColumnWidth = 30
$c.Columns.Item('C').ColumnWidth = 25
$c.Columns.Item('D').ColumnWidth = 15
$c.Columns.Item('E').ColumnWidth = 25

$c.Cells.Item(1,1) = "_provera Ime"
$c.Cells.Item(1,2) = "_provera Prezime"
$c.Cells.Item(1,3) = "_provera Filijala"
$c.Cells.Item(1,4) = "_AD Korisnik"
$c.Cells.Item(1,5) = "_AD Username"
$c.Cells.Item(1,6) = "_AD email"
$c.Cells.Item(1,7) = "_AD OU"
$c.Cells.Item(1,8) = "_AD upn"

#$c.Cells.Item(1,6) = "ParentContainerDN"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
#$d.EntireColumn.AutoFit($True)


$users = Import-Csv "D:\Scripts\_input\Kreiranje Prodavaca.csv"
#$kpp = "D:\Scripts\_Reports\Kreiranje Naloga\KreiraniNaloziZaProdavce.txt"

$i=2

foreach ($user in $users)

{
$FirstName = ($user.Ime).Trim()
$LastName = ($user.Prezime).Trim()
$desc = $user.Opis
$OUcsv = $user.Filijala.Trim()
#$jmbg = $user.JMBG


$Name = $FirstName+" "+$LastName
$lName = $LastName+" "+$FirstName

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

	Srem
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


#check user in OU container

$chk_user_ou = Get-QADUser -DisplayName $lName -SearchRoot $OU

if ($chk_user_ou) 
{
$c.Cells.Item($i,1) = $user.Ime
$c.Cells.Item($i,2) = $user.Prezime
$c.Cells.Item($i,3) = $user.Filijala
$c.Cells.Item($i,4) = $chk_user_ou.Name
$c.Cells.Item($i,5) = $chk_user_ou.samaccountname
$c.Cells.Item($i,6) = $chk_user_ou.email
$c.Cells.Item($i,7) = ($chk_user_ou.ParentContainer -replace "ddor.local/DDOR/","")
$c.Cells.Item($i,8) = $chk_user_ou.UserPrincipalName
}

elseif (Get-QADUser -DisplayName $lName )
{
$c.Cells.Item($i,1) = $user.Ime
$c.Cells.Item($i,2) = $user.Prezime
$c.Cells.Item($i,3) = $user.Filijala
$c.Cells.Item($i,4) = "Greska - user postoji ali u drugoj OU"
$c.Cells.Item($i,7) = ((Get-QADUser -DisplayName $lName ).ParentContainer -replace "ddor.local/DDOR/","")

}

Else 
{
$c.Cells.Item($i,1) = $user.Ime
$c.Cells.Item($i,2) = $user.Prezime
$c.Cells.Item($i,3) = $user.Filijala
$c.Cells.Item($i,4) = "Greska - nema naloga !!!"
}

$i++

}#end foreach