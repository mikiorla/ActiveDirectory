$ErrorActionPreference = "SilentlyContinue"

#$subnet = Read-Host "Subnet za pretragu DNS zapisa (npr. 206)"
#$ipsubnet = "192.168."+$subnet+".*"

$subnet = Read-Host "Subnet za pretragu DNS zapisa (npr. 100)"
$ipsubnet = "10.100."+$subnet+".*"


$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Name = "Serveri_"+$subnet
$c.Columns.Item('A').ColumnWidth = 28
$c.Columns.Item('B').ColumnWidth = 17
$c.Columns.Item('C').ColumnWidth = 15

$c.Cells.Item(1,1) = "Server Name"
$c.Cells.Item(1,2) = "IP"
$c.Cells.Item(1,3) = "TimeStamp"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)


$i=2

$dns = gwmi -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName adc 
 
$dns | Where-Object {$_.IPAddress -like $ipsubnet } | ForEach-Object {

if ($_.Timestamp -gt 0)
{
$ft = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1601 00:00').AddHours($_.Timestamp))
$t = Get-Date $ft -Format dd-MMM-yyyy 
}

else {

$t = $_.TimeStamp

}

$c.Cells.Item($i,1)=$_.ownername
$c.Cells.Item($i,2) = $_.IPaddress
$c.Cells.Item($i,3) = $t 
$i++

}

