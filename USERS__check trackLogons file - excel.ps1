
# reads trackLogons.txt file where useename and computer name are loged
# check DNS for IP 
# writes to excel



$ErrorActionPreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Columns.Item('A').ColumnWidth = 15
$c.Columns.Item('B').ColumnWidth = 30
$c.Columns.Item('C').ColumnWidth = 25
#$c.Columns.Item('D').ColumnWidth = 15
#$c.Columns.Item('E').ColumnWidth = 25

$c.Cells.Item(1,1) = "Username"
$c.Cells.Item(1,2) = "Comp"
$c.Cells.Item(1,3) = "IP"
#$c.Cells.Item(1,4) = "Company"
#$c.Cells.Item(1,5) = "Department"
#$c.Cells.Item(1,6) = "Office"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)

#get DNS records
$dns = gwmi -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName novisad2
$a = Get-Content \\vm-corefs\winadmins$\tracklogons.txt
$i=2
foreach ($pair in $a)
{

$p = $pair.Split()

$username = $p[0]
$comp = $p[1]

$c.Cells.Item($i,1) = $username
$c.Cells.Item($i,2) = $comp

$compFQDN = $comp+".ddor.local"

#check DNS for comp IP
$dns | ? {$_.OwnerName -eq $compFQDN} | Set-Variable compDNS

$c.Cells.Item($i,3) = $compDNS.IPAddress


$i++
}