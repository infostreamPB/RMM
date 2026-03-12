# -------------------------------
# Robust Offline RMM Cleanup Script
# -------------------------------

# ITSPlatform Product GUID
$ITSPlatformGUID = "{18f39771-f9d8-4cfd-9654-f6c67c8ad9f4}"

function Uninstall-ITSPlatform {
    Write-Host "Attempting to uninstall ITSPlatform..."

    # Option 1: WMIC
    if (Get-Command wmic -ErrorAction SilentlyContinue) {
        Write-Host "Using WMIC to uninstall ITSPlatform..."
        wmic product where "name like '%ITSPlatform%'" call uninstall /nointeractive
        return
    }

    # Option 2: PowerShell PackageManagement
    if (Get-Command Get-Package -ErrorAction SilentlyContinue) {
        $pkg = Get-Package | Where-Object { $_.Name -like '*ITSPlatform*' }
        if ($pkg) {
            Write-Host "Using PowerShell PackageManagement to uninstall ITSPlatform..."
            foreach ($p in $pkg) {
                Uninstall-Package -Name $p.Name -Force -ErrorAction SilentlyContinue
            }
            return
        }
    }

    # Option 3: MSIExec
    if (Test-Path "C:\Windows\System32\msiexec.exe") {
        Write-Host "Using MSIExec to uninstall ITSPlatform..."
        Start-Process msiexec.exe -ArgumentList "/x $ITSPlatformGUID /qn /norestart" -Wait
        return
    }

    Write-Warning "Unable to uninstall ITSPlatform using WMIC, PackageManagement, or MSIExec. Will remove folders manually."
}

# Run uninstall
Uninstall-ITSPlatform

# Stop services safely
$services = @(
    "SAAZappr",
    "SAAZDPMACTL",
    "SAAZRemoteSupport",
    "SAAZScheduler",
    "SAAZServerPlus",
    "SAAZWatchDog"
)

foreach ($svc in $services) {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        Write-Host "Stopping service $svc..."
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    }
}

# Remove program folders safely
$folders = @(
    "C:\Program Files (x86)\SAAZOD",
    "C:\Program Files (x86)\SAAZODBKP"
)

foreach ($folder in $folders) {
    if (Test-Path $folder) {
        Write-Host "Removing folder $folder..."
        Remove-Item -Path $folder -Force -Recurse -ErrorAction SilentlyContinue
    }
}

# Remove registry properties safely
$regProps = @(
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest"; Name = "ITSPlatformID" }
)

foreach ($prop in $regProps) {
    if (Get-ItemProperty -Path $prop.Path -Name $prop.Name -ErrorAction SilentlyContinue) {
        Write-Host "Removing registry property $($prop.Name)..."
        Remove-ItemProperty -Path $prop.Path -Name $prop.Name -Force -ErrorAction SilentlyContinue
    }
}

# Remove registry keys safely
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

Write-Host "`nRMM cleanup complete."
