Function WMILookup 
{
$AllServers = @()
$ServerObj  = @()
foreach ($StrComputer in $Computers)
{
$NetItems = $null
Write-Output “Working on $StrComputer”
$ServerObj = New-Object psObject
$ServerObj | Add-Member -membertype noteproperty -name “Hostname” -Value  $StrComputer -Force
$NetItems = @(gwmi Win32_NetworkAdapterConfiguration -Comp $StrComputer | where{$_.IPEnabled -eq “True”} | select @{name=‘IP Address’;expression={$_.IPAddress}},MACAddress,DNSServerSearchOrder)
$intRowNet = 0
$ServerObj | Add-Member -membertype noteproperty -name “NIC’s” -Value $NetItems.Length -Force
[STRING]$MACAddresses = @()
[STRING]$IpAddresses = @()
[STRING]$DNS = @()
foreach ($objItem in $NetItems)
{
if ($objItem.{IP Address}.Count -gt 1)
{
$TempIpAdderesses = [STRING]$objItem.{IP Address}
$TempIpAdderesses  = $TempIpAdderesses.Trim().Replace(” “, ” ; “)
$IpAddresses += $TempIpAdderesses # +”; “
}
else
{
$IpAddresses += $objItem.{IP Address} +“; “
}
if ($objItem.{MacAddress}.Count -gt 1)
{
$TempMACAddresses = [STRING]$objItem.{MACAddress}
$TempMACAddresses = $TempMACAddresses.Replace(” “, ” ; “)
$MACAddresses += $TempMACAddresses +“; “
}
else
{
$MACAddresses += $objItem.{MACAddress} +“; “
}
if ($objItem.{DNSServerSearchOrder}.Count -gt 1)
{
$TempDNSAddresses = [STRING]$objItem.{DNSServerSearchOrder}
$TempDNSAddresses = $TempDNSAddresses.Replace(” “, ” ; “)
$DNS += $TempDNSAddresses +“; “
}
else
{
$DNS += $objItem.{DNSServerSearchOrder} +“; “
}
$intRowNet = $intRowNet + 1
}
$ServerObj | Add-Member -membertype noteproperty -name “IP Address” -Value $IpAddresses.substring(0,$ipaddresses.LastIndexOf(“;”)) -Force
$ServerObj | Add-Member -membertype noteproperty -name “MAC Address” -Value $MACAddresses.substring(0,$MACAddresses.LastIndexOf(“;”)) -Force
$ServerObj | Add-Member -MemberType NoteProperty -Name “Home OU” -Value $ComputerOU[$strComputer] -Force
$ServerObj | Add-Member -MemberType NoteProperty -Name “Function” -Value $ComputerDescription[$strComputer] -Force
$ServerObj | Add-Member -MemberType NoteProperty -Name “DNS” -Value $DNS
$AllServers += $ServerObj
}
$file = “ServerDNS.csv”
$AllServers |Sort-Object “HostName” | Export-Csv “$file” -NoTypeInformation
}
Function ListADServers 
{
$Computers = @()
$DomainName = [ADSI]”
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
$objSearcher.SearchRoot = $DomainName 
$objSearcher.SearchRoot = $Root
$objSearcher.SearchScope = “SubTree”
$objSearcher.PageSize = 1000
$objSearcher.PropertiesToLoad.Add(“adspath”)
$objSearcher.Filter =  “(&(objectCategory=computer)(OperatingSystem=Windows*Server*))”
$colResults = $objSearcher.FindAll()
foreach ($objResult in $colResults)
{
$Computers += $objResult.Properties.adspath
}
}
$erroractionpreference = “SilentlyContinue”
. ListADServers
WMILookup