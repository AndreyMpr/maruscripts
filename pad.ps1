Start-Process "C:\Program Files (x86)\Power Automate Desktop\PAD.Console.Host.exe"
Start-Sleep -Seconds 5
$wshell = New-Object -ComObject wscript.shell;
$wshell.SendKeys("^%+8")