#Uninstall WMF 5 Exchange
$HotFixToUninstall = Get-HotFix -Id KB3066437
                 if ($HotFixToUninstall) {
                    Start-Process -FilePath "c:\windows\system32\wusa.exe" -ArgumentList "/uninstall /kb:3066437 /quiet /norestart" 
                 } 