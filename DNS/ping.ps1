$readfile=get-content "D:\ScriptsFinal\_PS\dns\Hosts.txt" 

foreach($readf in $readfile) {
 
$ALive=get-wmiobject win32_pingstatus -Filter "Address='$readf'" | Select-Object statuscode

	if($ALive.statuscode -eq 0) 

		{write-host $readf is REACHABLE -foreground "GrEEN"} 
	else 
		{write-host $readf is NOT reachable -foreground "RED"} 
}