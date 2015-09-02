$ErrorActionPreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Columns.Item('A').ColumnWidth = 15
$c.Columns.Item('B').ColumnWidth = 30
$c.Columns.Item('C').ColumnWidth = 25
$c.Columns.Item('D').ColumnWidth = 15
$c.Columns.Item('E').ColumnWidth = 25

#$c.Cells.Item(1,1) = "eMail"
$c.Cells.Item(1,1) = "Input"
$c.Cells.Item(1,2) = "SAM"
$c.Cells.Item(1,3) = "DisplayName"
#$c.Cells.Item(1,4) = "StoreName"
$c.Cells.Item(1,4) = "Company"
$c.Cells.Item(1,5) = "_new Company"

$c.Cells.Item(1,6) = "Department"
$c.Cells.Item(1,7) = "_new Department"
$c.Cells.Item(1,8) = "Office"
$c.Cells.Item(1,9) = "_new Office"
#$c.Cells.Item(1,7) = "PasswordLastSet"
#$c.Cells.Item(1,8) = "ADstatus"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)


#$emails = Get-Content D:\ScriptsFinal\_PS\AD\Users\prodajnamreza-emails.txt 
#$emails = Get-Content D:\ScriptsFinal\_PS\AD\Users\CSV\pop3-mailboxes.txt
#$emails = Get-Content D:\ScriptsFinal\_PS\AD\Users\prodajnamreza-emails.txt
#$prodajna = [ADSI]"LDAP://CN=Prodajna Mreza - Sales Network,OU=_Distributivne Liste,OU=DDOR,DC=ddor,DC=local"
$i = 2
#$sams = Get-content "D:\ScriptsFinal\_PS\AD\Users\zaTehnickeS.txt"
#$sams = Get-QADUser -SearchRoot "OU=DDOR,DC=ddor,DC=local" -Email *
#$sams = get-qaduser -searchroot "ddor.local/DDOR Administracija/_Administrativni nalozi i grupe/Domain admins nalozi" -Email *

#$sams = get-qaduser -searchroot "ddor.local/DDOR" -Email * -SizeLimit 0
#$sams = get-qaduser -searchroot "ddor.local/DDOR" -SizeLimit 0


#$sams = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika CENTRALA 27Sep2011.txt' -Delimiter ";"
#$sams = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika SteteOU 28Sep2011.txt' -Delimiter ";"
#$sams = Import-Csv 'd:\Scripts\_input\Users\Setovanje korisnika oNS1 06Okt11.txt' -Delimiter ";"
#$sams = Import-Csv 'd:\Scripts\_input\Users\Setovanje korisnika Region2 6-Okt-2011.txt' -Delimiter ";"
#$sams = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Region345 6-Okt-2011.txt' -Delimiter ";"
#$sams = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Region 678.txt' -Delimiter ";"
$sams = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Sektor prodaje.txt' -Delimiter ";"

#$sams = Get-Content 'H:\Ad test UAC.txt'

foreach ($sam_u in $sams)
{
	
	#Write-Host $sam_u.SamAccountName
	#$chk_u = Get-QADUser -SamAccountName $sam_u -IncludedProperties userAccountControl
	#$chk_u = Get-QADUser $sam_u -IncludedProperties userAccountControl
	
	$chk_u = Get-QADUser -DisplayName $sam_u.punoime
	
	if ($chk_u) 

	{
				
		$c.Cells.Item($i,1) = $sam_u.PunoIme	
	  	$c.Cells.Item($i,2) = $chk_u.samaccountname
		#$c.Cells.Item($i,4) = $chk_u.userAccountControl
		$c.Cells.Item($i,3) = $chk_u.DisplayName
		$c.Cells.Item($i,4) = $chk_u.Company
		$c.Cells.Item($i,5) = $sam_u.kompanija	
		#$c.Cells.Item($i,5) = $chk_u.LastName
		#$c.Cells.Item($i,5) = $chk_u.mail
		#$c.Cells.Item($i,5) = $chk_u.ParentContainerDN
		$c.Cells.Item($i,6) = $chk_u.Department
		$c.Cells.Item($i,7) = $sam_u.departman
		#$c.Cells.Item($i,6) = $chk_u.UserMustChangePassword
		$c.Cells.Item($i,8) = $chk_u.Office
		$c.Cells.Item($i,9) = $sam_u.kancelarija	
		#$c.Cells.Item($i,7) = $chk_u.PAsswordLastSet
		
		#if ($chk_u.AccountIsDisabled) {$c.Cells.Item($i,8) = "Disabled"} else {$c.Cells.Item($i,8) = "Enabled"}
		
		$i++
		 	
	}
	else {
	$c.Cells.Item($i,1) = $sam_u.PunoIme
	$c.Cells.Item($i,2) = "Nema naloga na mrezi"
	$i++ }

	
}


