$ErrorActionPreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True
$b = $a.Workbooks.Add()

$msdns_class = gwmi -ComputerName novisad2 -Namespace root\MicrosoftDNS -List | ? {$_.name -like "*MicrosoftDNS*"}
foreach ($srv_type in $msdns_class)

{

$c = $b.Worksheets.Add()
$c.Name = $srv_type.Name
$c.Columns.Item('A').ColumnWidth = 28
$c.Columns.Item('B').ColumnWidth = 17
$c.Columns.Item('C').ColumnWidth = 15

$c.Cells.Item(1,1) = "Server Name"
$c.Cells.Item(1,2) = "RecordData"
$c.Cells.Item(1,3) = "TextRepresentation"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)

$i=2

$dns = gwmi -Namespace root\MicrosoftDNS -Class $srv_type.Name -ComputerName novisad2
foreach ($record in $dns)
{

$c.Cells.Item($i,1)= $record.DomainName
$c.Cells.Item($i,2) = $record.RecordData
$c.Cells.Item($i,3) = $record.TextRepresentation 
$i++

}



}