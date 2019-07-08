Set oShell = CreateObject ("Wscript.Shell")
oShell.run "cmd.exe /C netsh wlan show profiles & PAUSE"
Dim strName
Wscript.sleep 200
strName = InputBox("Enter name of wireless network: ")
oShell.run "cmd.exe /C netsh wlan show profile name=""" & strName & """ key=clear | findstr Key & PAUSE"
Wscript.sleep 4500
oShell.SendKeys "{ENTER}"
Wscript.sleep 200
oShell.SendKeys "{ENTER}"