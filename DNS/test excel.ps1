
$ErrorActionPreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True
$b = $a.Workbooks.Add()

$items = 1..10

foreach ($x in $items)
{
$i=2
$c = $b.Worksheets.Add()
$c.Name = $x
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



#$dns = gwmi -Namespace root\MicrosoftDNS -Class $srv_type.Name -ComputerName novisad2



$c.Cells.Item($i,1)= "test1"
$c.Cells.Item($i,2) = "test2"
$c.Cells.Item($i,3) = "test3" 
$i++


}