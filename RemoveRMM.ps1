# Force TLS 1.2 for any web actions (optional if you download scripts)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Uninstall software safely
try {
    $apps = Get-Package -Name "*ITSPlatform*" -ErrorAction SilentlyContinue
    foreach ($app in $apps) {
        Write-Host "Uninstalling $($app.Name)..."
        Uninstall-Package $app -Force -ErrorAction SilentlyContinue
    }
} catch {
    Write-Host "No ITSPlatform package found, skipping."
}

# Stop services safely
$services = "SAAZappr","SAAZDPMACTL","SAAZRemoteSupport","SAAZScheduler","SAAZServerPlus","SAAZWatchDog"
foreach ($svc in $services) {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Write-Host "Stopping service $svc..."
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    }
}

# Remove program folders safely
$folders = "C:\Program Files (x86)\SAAZOD","C:\Program Files (x86)\SAAZODBKP"
foreach ($folder in $folders) {
    if (Test-Path $folder) {
        Write-Host "Removing folder $folder..."
        Remove-Item -Path $folder -Force -Recurse -ErrorAction SilentlyContinue
    }
}

# Remove registry properties and keys safely
$regProps = @(
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest"; Name = "ITSPlatformID" }
)
foreach ($prop in $regProps) {
    if (Get-ItemProperty -Path $prop.Path -Name $prop.Name -ErrorAction SilentlyContinue) {
        Write-Host "Removing registry property $($prop.Name)..."
        Remove-ItemProperty -Path $prop.Path -Name $prop.Name -Force -ErrorAction SilentlyContinue
    }
}

$regKeys = @(
    "HKLM:\SOFTWARE\WOW6432Node\SAAZOD",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZappr",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZDPMACTL",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZRemoteSupport",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZScheduler",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZServerPlus",
    "HKLM:\SYSTEM\CurrentControlSet\Services\SAAZWatchDog"
)
foreach ($key in $regKeys) {
    if (Test-Path $key) {
        Write-Host "Removing registry key $key..."
        Remove-Item -Path $key -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Write-Host "RMM cleanup complete."
