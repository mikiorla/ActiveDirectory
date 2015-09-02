#
#Setovanje atributa korisnickih naloga iz csv fajla
#

$greske = New-Object System.Collections.ArrayList
$datum_log = [string](get-date) -replace "/","-" -replace " ","_" -replace ":","-"
$logname = "setovanje_atributa_"+$datum_log+".txt"
New-Item -ItemType File -Path "D:\Scripts\_Reports\Users" -Name $logname | Set-Variable log

#eSlekcija korisnika iz CSV fajla

#$users = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika CENTRALA 27Sep2011.txt' -Delimiter ";"
#$users = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika SteteOU 28Sep2011.txt' -Delimiter ";"
#$users = Import-Csv 'd:\Scripts\_input\Users\Setovanje korisnika oNS1 06Okt11.txt' -Delimiter ";"
#$users = Import-Csv 'd:\Scripts\_input\Users\Setovanje korisnika Region2 6-Okt-2011.txt' -Delimiter ";"
#$users = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Region345 6-Okt-2011.txt' -Delimiter ";"
#$users = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Region 678.txt' -Delimiter ";"
#$users = Import-Csv 'D:\Scripts\_input\Users\Setovanje korisnika Sektor prodaje.txt' -Delimiter ";"
$users = import-csv 'D:\Scripts\_input\Users\preostali korisnici.txt' -Delimiter ";"

$date = Get-Date
$o = "___Setovanje korisnika___"+ $datum_log
Write-Output $o >> $log
Write-Output `n >> $log
Write-Host $o

foreach ($user in $users)
{
#$display_name = $user.PunoIme
$SAM = $user.PunoIme
$new_company = $user.kompanija
$new_department = $user.departman
$new_office = $user.kancelarija

#$sam = $user.SamAccountIme
#$displayN = $user.DisplayIme
#$desc = $user.Posao
#$mesto = $user.MestoRada

$ad_user = Get-QADUser -samaccountname $SAM

#if (($ad_user) -and ($ad_user.Count -eq 1))
if ($ad_user)
{
$o1 = $ad_user.DisplayName+" "+$ad_user.SamAccountName
#Write-Output $o1 >> $log
#Write-Host $o1 ":" $new_title $new_office
#$o11 = " (old_title:" + $ad_user.title + ", new:"+$new_title+")"
$o12 = " Old office:"+ $ad_user.office+", new:"+$new_office
$o13 = " Old Department:"+ $ad_user.department+", new:"+$new_department
$o111 = $o1 + $o11 + $o12+$o13
Write-Output $o111 >> $log

# Akcija 1 (set Description)
Set-QADUser $ad_user -Company $new_company -Department $new_department -Office $new_office



}

Else 
{
$o2 = "Nema naloga na AD-u ili Greska - "+$user.PunoIme
#Write-Output $o2 >> $log 
Write-Host $o2
$display_name_x = $display_name+"`n"
$null_p = $greske.Add($display_name_x)


}


$o2,$o1 = $null
}

Write-Output $greske >> $log
Write-Output `n >> $log
Write-Host $greske