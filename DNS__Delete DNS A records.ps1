

$rr = Get-Content 'D:\ScriptsFinal\_PS\dns\input\DNS A records to delete.txt'
$reportname = "DNS deleted records "+(Get-Date -Format dd-MMMM-yyyy_hh-mm-ss)+".txt"
$report = "D:\ScriptsFinal\_PS\dns\reports\"+$reportname
Write-output "Obrisani DNS zapisi:" > $report
Write-output "====================" >> $report
Write-Host "DNS record to delete..."
foreach ($rr_a in $rr)
{
$dnsrr = gwmi -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName novisad2 -Filter "OwnerName='$rr_a'"
$dnsrr.psbase.Delete() 
$logging = $dnsrr.OwnerName+" "+$dnsrr.IPAddress
Write-Output  $logging >> $report
Write-host "DNS record deleted..."$dnsrr.OwnerName
}


#***** Delete Resorce Record *****#
#$rec = Get-WmiObject -ComputerName DC02 -Namespace 'root\MicrosoftDNS' -Class MicrosoftDNS_AType  -Filter "IPAddress = '192.168.172.20'"
#$rec.psbase.Delete()


#***** uncheck 'Delete this record when it becomes stale' on DNS A record *****#
#$record.timeStamp = 0 # checkbox is unchecked,  3579756 is checked
#$record.psbase.put()
