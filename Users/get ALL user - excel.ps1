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
#$c.Cells.Item(1,1) = "Korisnik"
$c.Cells.Item(1,2) = "SamAccountName"
$c.Cells.Item(1,3) = "DisplayName"
#$c.Cells.Item(1,4) = "StoreName"
#$c.Cells.Item(1,3) = "Office"
#$c.Cells.Item(1,3) = "Office"
#$c.Cells.Item(1,4) = "OU"
#$c.Cells.Item(1,5) = "PasswordAge"
#$c.Cells.Item(1,6) = "AccountIsDisabled"
#$c.Cells.Item(1,7) = "Department"
#$c.Cells.Item(1,6) = "Company"
#$c.Cells.Item(1,7) = "Title"
#$c.Cells.Item(1,8) = "ADstatus"
$c.Cells.Item(1,4) = "Email"
$c.Cells.Item(1,1) = "SAP_APL"


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

#$sams = get-qaduser -searchroot "ddor.local/DDOR" -SizeLimit 0 | ? {!($_.AccountIsDisabled)}
#$sams = get-qaduser -searchroot "ddor.local/DDOR" -SizeLimit 0 
$sams = Get-Content D:\Temp\sap_korisnici.txt

#$sams = get-qaduser -searchroot "ddor.local/DDOR" -SizeLimit 0 -includeallproperties 
#$sams = get-qaduser -searchroot "ddor.local/DDOR" -SizeLimit 0 -PasswordNotChangedFor 120 | ? {!($_.AccountIsDisabled)}

foreach ($sam_u in $sams)
{
	#Write-Host $sam_u.SamAccountName
	#$chk_u = Get-QADUser -samaccountname $sam_u.samaccountname
	$chk_u = Get-QADUser -samaccountname $sam_u
	
	if ($chk_u) 

	{
				
		$c.Cells.Item($i,2) = $chk_u.samaccountname
	  	$c.Cells.Item($i,3) = $chk_u.DisplayName
		#$c.Cells.Item($i,3) = $chk_u.displayname
		#$c.Cells.Item($i,3) = $chk_u.Office
		#$c.Cells.Item($i,4) = $chk_u.ParentContainerDN
		#$c.Cells.Item($i,5) = [string]$chk_u.PasswordAge
		#$c.Cells.Item($i,6) = $chk_u.AccountIsDisabled
		#$c.Cells.Item($i,7) = $chk_u.Department
		$c.Cells.Item($i,4) = $chk_u.email
		$c.Cells.Item($i,1) = $sam_u
		#$c.Cells.Item($i,7) = $chk_u.ModificationDate
		#$c.Cells.Item($i,7) = $chk_u.Title
		#$c.Cells.Item($i,8) = $chk_u.Description
		#if ($chk_u.AccountIsDisabled) {$c.Cells.Item($i,8) = "Disabled"} else {$c.Cells.Item($i,8) = "Enabled"}
		
		$i++
		 	
	}
	else {
	$c.Cells.Item($i,1) = $sam_u
	$c.Cells.Item($i,2) = "Nema naloga na mrezi"
	$i++ }

	
}


