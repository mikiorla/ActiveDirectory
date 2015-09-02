$readfile=get-content "D:\ScriptsFinal\_PS\dns\Hosts.txt" 

foreach($readf in $readfile) {
 
(New-Object -comobject Wscript.Shell).Run("cmd /C ipconfig /all >> D:\ScriptsFinal\_PS\dns\ipconfig.txt") 
sleep 10

}