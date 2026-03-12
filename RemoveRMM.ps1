[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 wmic product where "name like '%ITSPlatform%'" call uninstall /nointeractive


Stop-Service -Name "SAAZappr"
Stop-Service -Name "SAAZDPMACTL"
Stop-Service -Name "SAAZRemoteSupport"
Stop-Service -Name "SAAZScheduler"
Stop-Service -Name "SAAZServerPlus"
Stop-Service -Name "SAAZWatchDog"
If (Test-Path "C:\Program Files (x86)\SAAZOD"){
   Remove-Item "C:\Program Files (x86)\SAAZOD" -Force -Recurse
} else {}
If (Test-Path "C:\Program Files (x86)\SAAZODBKP"){
   Remove-Item "C:\Program Files (x86)\SAAZODBKP" -Force -Recurse
} else {}
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest" -Name "ITSPlatformID" -Force
Remove-Item "HKLM:\SOFTWARE\WOW6432Node\SAAZOD" -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZappr" -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZDPMACTL" -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZRemoteSupport" -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZScheduler" -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZServerPlus" -Force

Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZWatchDog" -Force
