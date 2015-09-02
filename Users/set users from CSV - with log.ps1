#
#Kreiranje korisnickih naloga i mailboxova iz csv fajla
#
$datum = [string](get-date) -replace "/","-" -replace " ","_" -replace ":","-"

$log = "D:\Scripts\_Reports\Users\Setovanje Korisnika.txt"

#Slekcija korisnika iz CSV fajla
#$users = Import-Csv "D:\Scripts\_input\Users\Becej korisnici csv.txt" -Delimiter ","
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\BackaPalanka korisnici csv.txt" 
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\BackaTopola korisnici csv.txt" 
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\Subotica korisnici csv.txt" 
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\Jagodina korisnici csv.txt"
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\kikinda korisnici csv.txt"
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\kragujevac korisnici csv.txt"
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\kula-nis-novipazar korisnici csv.txt"
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\novisad korisnici csv.txt"
#$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\pa-ru-se-si-sd-so-sm-sp korisnici csv.txt"
$users = Import-Csv -Delimiter ";" "D:\Scripts\_input\Users\ue-va-vr-vs-zr korisnici csv.txt"
$date = Get-Date
$o = "___Setovanje Korisnika___ue-va-vr-vs-zr"+ $datum
Write-Output $o >> $log
Write-Output `n >> $log
Write-Host $o

foreach ($user in $users)
{
$sam = $user.SamAccountIme
$displayN = $user.DisplayIme
$desc = $user.Posao
$mesto = $user.MestoRada

$ad_user = Get-QADUser -SamAccountName $sam

if ($ad_user) 
{
$o1 = $ad_user.SamAccountName + " (" + $ad_user.DisplayName + ")"
#Write-Output $o1 >> $log
Write-Host $o1 ":" $desc
$o11 = " (stari Description:" + $ad_user.Description + ", novi:"+$desc+")"
$o111 = $o1 + $o11
Write-Output $o111 >> $log
# Akcija 1 (set Description)
Set-QADUser $ad_user -Description $desc | Out-Null 

}

Else 
{
$o2 = "Korisnik nije pronadjen na AD-u:"+$user
Write-Output $o2 >> $log 
Write-Host $o2
}


$o2,$o1 = $null
}
Write-Output `n >> $log