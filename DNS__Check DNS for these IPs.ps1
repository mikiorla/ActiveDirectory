#
$ips = Get-Content "D:\ScriptsFinal\_PS\dns\input\Check DNS for these IP.txt"
#$dnsrr = gwmi -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName novisad2 
foreach ($ip in $ips)
{
$ipss = $ip.Trim()
$dnsrr = gwmi -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName novisad2 -Filter "IPAddress='$ipss'"
$aa = $dnsrr.OwnerName
if($aa) {Write-host $ip$aa}
}
