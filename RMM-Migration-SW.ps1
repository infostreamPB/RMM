# ==========================================
# RMM MIGRATION (EXACT FILENAME PARSING)
# ==========================================
$ErrorActionPreference = "Stop"

# =============================
# CONFIG
# =============================
$TaskName   = "RMM-Migration-Job"
$WorkDir    = "C:\ProgramData\RMM-Migration"
$ScriptPath = "$WorkDir\runner.ps1"
$LogFile    = "$WorkDir\migration.log"
$Transcript = "$WorkDir\migration-transcript.log"
# =====================================================================
# REPLACE THE URL BELOW WITH YOUR NEW AGENT DOWNLOAD URL
# Get this URL from the ConnectWise Automate/ScreenConnect RMM portal
# =====================================================================
$NewAgentUrl  = "https://prod.setup.itsupport247.net/windows/BareboneAgent/32/West_Palm_Beach-Simmons_%26_White_Windows_OS_ITSPlatform_TKN8064b9d0-9c4c-45c9-86cc-6a1641e77e84/MSI/setup"
# =====================================================================
# DO NOT MODIFY ANYTHING BELOW THIS LINE
# =====================================================================
$NewAgentArgs = "/quiet /norestart ALLUSERS=1"
$NewAgentInstallPath = "C:\Program Files (x86)\ITSPlatform"
$UninstallExe = "C:\PROGRA~2\SAAZOD\Uninstall\uninstall.exe"

# Clean up any previous run
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue }

if (!(Test-Path $WorkDir)) { New-Item $WorkDir -ItemType Directory -Force | Out-Null }

# =============================
# RUNNER SCRIPT TEMPLATE
# =============================
$RunnerTemplate = @'
$WorkDir    = "{0}"
$LogFile    = "{1}"
$Transcript = "{2}"
$AgentUrl   = "{3}"
$AgentArgs  = "{4}"
$InstallPath = "{5}"
$OldUninst  = "{6}"
$TaskName   = "{7}"

function Write-Log {{
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Msg"
    try {{ $line | Out-File -FilePath $LogFile -Append -Encoding UTF8 }} catch {{}}
    Write-Output $line
}}

try {{ Start-Transcript -Path $Transcript -Append -Force | Out-Null }} catch {{
    Write-Log "Start-Transcript failed: $($_.Exception.Message)" "WARN"
}}

$ErrorActionPreference = 'Stop'
Write-Log "=== Migration Started ==="
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "OS: $([System.Environment]::OSVersion.VersionString)"
Write-Log "PowerShell: $($PSVersionTable.PSVersion)"
Write-Log "User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# Force TLS 1.2 (required for Windows Server 2012 R2 where it is not the default)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- DOWNLOAD ---
try {{
    Write-Log "Downloading agent from: $AgentUrl"
    $response = Invoke-WebRequest -Uri $AgentUrl -UseBasicParsing -ErrorAction Stop

    # Get the original filename from Content-Disposition header
    $cd = $response.Headers['Content-Disposition']
    if ($cd -and $cd -match 'filename\s*=\s*"?([^";]+)"?') {{
        $originalFileName = [uri]::UnescapeDataString($Matches[1].Trim())
    }} elseif ($response.BaseResponse.ResponseUri) {{
        $originalFileName = [uri]::UnescapeDataString($response.BaseResponse.ResponseUri.Segments[-1].TrimEnd('/'))
    }} else {{
        $originalFileName = [uri]::UnescapeDataString(([uri]$AgentUrl).Segments[-1].TrimEnd('/'))
    }}

    if ($originalFileName -and -not $originalFileName.EndsWith('.msi')) {{
        $originalFileName = $originalFileName + ".msi"
    }}

    $agentPath = Join-Path $WorkDir $originalFileName
    [IO.File]::WriteAllBytes($agentPath, $response.Content)
    Write-Log "File saved with original name: $originalFileName"
    Write-Log "Full path: $agentPath"
    Write-Log "File size: $((Get-Item $agentPath).Length) bytes"
}} catch {{
    Write-Log "Download FAILED: $($_.Exception.Message)" "ERROR"
    exit 1
}}

# =====================================================================
# CLEANUP PHASE
# =====================================================================
Write-Log "=== Beginning Full Cleanup ==="

# --- STEP 1: Stop ITSPlatform services ---
Write-Log "Stopping ITSPlatform services..."
Stop-Service -Name "ITSPlatform" -Force -ErrorAction SilentlyContinue
Stop-Service -Name "ITSPlatformManager" -Force -ErrorAction SilentlyContinue
Start-Sleep 3

# --- STEP 2: Kill ITSPlatform processes by name ---
$itsProcesses = @("platform-communicator-tray", "platform-communicator-plugin", "platform-agent-core", "platform-agent-manager", "platform-eventlog-plugin", "platform-sysevents-plugin")
foreach ($proc in $itsProcesses) {{
    try {{
        $p = Get-Process -Name $proc -ErrorAction SilentlyContinue
        if ($p) {{
            Stop-Process -Name $proc -Force
            Write-Log "Killed process: $proc"
        }}
    }} catch {{
        Write-Log "Error stopping process $proc : $($_.Exception.Message)" "WARN"
    }}
}}

# --- STEP 3: Uninstall ITSPlatform via package manager ---
try {{
    $itsPlatformEntry32 = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object {{ $_.DisplayName -eq "ITSPlatform" }}
    $itsPlatformEntry64 = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object {{ $_.DisplayName -eq "ITSPlatform" }}
    if ($itsPlatformEntry32 -or $itsPlatformEntry64) {{
        Write-Log "Uninstalling ITSPlatform via Get-Package..."
        Get-Package -Name "ITSPlatform*" -ErrorAction SilentlyContinue | Uninstall-Package -Force -ErrorAction SilentlyContinue
        Start-Sleep 5
    }}
}} catch {{
    Write-Log "Error uninstalling ITSPlatform package: $($_.Exception.Message)" "WARN"
}}

# --- STEP 4: Stop and kill SAAZOD processes ---
$saazProcesses = @("SAAZappr", "SAAZDPMACTL", "SAAZRemoteSupport", "SAAZScheduler", "SAAZServerPlus", "SAAZWatchDog")
foreach ($proc in $saazProcesses) {{
    try {{
        $p = Get-Process -Name $proc -ErrorAction SilentlyContinue
        if ($p) {{
            Stop-Process -Name $proc -Force
            Write-Log "Killed SAAZOD process: $proc"
        }}
    }} catch {{
        Write-Log "Error stopping SAAZOD process $proc : $($_.Exception.Message)" "WARN"
    }}
}}

# --- STEP 5: Delete SAAZOD services ---
foreach ($svc in $saazProcesses) {{
    $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($service) {{
        Write-Log "Deleting SAAZOD service: $svc"
        sc.exe delete $svc | Out-Null
        Start-Sleep 2
    }}
}}

# --- STEP 6: Run SAAZOD uninstaller with proper arguments ---
if (Test-Path $OldUninst) {{
    Write-Log "Running legacy SAAZOD uninstaller..."
    Start-Process -FilePath $OldUninst -ArgumentList "/U:C:\PROGRA~2\SAAZOD\Uninstall\uninstall.xml", "/silent" -Wait -ErrorAction SilentlyContinue
    Start-Sleep 5
}}

# --- STEP 7: Delete SAAZOD registry key ---
Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\SAAZOD" -Recurse -Force -ErrorAction SilentlyContinue
Write-Log "Removed SAAZOD registry key"

# --- STEP 8: Delete ITSPlatform/SAAZOD uninstall registry entries ---
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)
foreach ($path in $registryPaths) {{
    Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {{
        $displayName = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
        if ($displayName -match "SAAZOD|ITSPlatform") {{
            Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Removed uninstall registry entry: $displayName"
        }}
    }}
}}

# --- STEP 9: Delete Installer\Folders registry entries ---
$folderRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders"
$folderEntries = @(
    "C:\Program Files (x86)\ITSPlatform\", "C:\Program Files (x86)\ITSPlatform\legacysetup\",
    "C:\Program Files (x86)\ITSPlatform\db\", "C:\Program Files (x86)\ITSPlatform\log\",
    "C:\Program Files (x86)\ITSPlatform\tmp\", "C:\Program Files (x86)\ITSPlatform\agentcore\",
    "C:\Program Files (x86)\ITSPlatform\installationmanager\", "C:\Program Files (x86)\ITSPlatform\agentmanager\",
    "C:\Program Files (x86)\ITSPlatform\config\", "C:\Program Files (x86)\ITSPlatform\delta_config\",
    "C:\Program Files (x86)\ITSPlatform\delta_config\brightgauge\", "C:\Program Files (x86)\ITSPlatform\delta_config\dt\",
    "C:\Program Files (x86)\ITSPlatform\delta_config\integration\", "C:\Program Files (x86)\ITSPlatform\delta_config\production\",
    "C:\Program Files (x86)\ITSPlatform\delta_config\qa\", "C:\Program Files (x86)\ITSPlatform\delta_config\remove_schedule\",
    "C:\Program Files (x86)\ITSPlatform\delta_config\sandbox\", "C:\Program Files (x86)\ITSPlatform\delta_config\sandboxdev\",
    "C:\Program Files (x86)\ITSPlatform\delta_config\staging\", "C:\Program Files (x86)\ITSPlatform\plugin\",
    "C:\Program Files (x86)\ITSPlatform\plugin\asset\", "C:\Program Files (x86)\ITSPlatform\plugin\version\",
    "C:\Program Files (x86)\ITSPlatformSetupLogs\", "C:\Program Files (x86)\ITSPlatformSetupLogs\logbackup\"
)
foreach ($entry in $folderEntries) {{
    Remove-ItemProperty -Path $folderRegPath -Name $entry -ErrorAction SilentlyContinue
}}
Write-Log "Cleaned Installer\Folders registry entries"

# --- STEP 10: Delete ITSPlatform service registry entries + WMI fallback ---
$servicesToDelete = @("ITSPlatform", "ITSPlatformManager")
foreach ($svc in $servicesToDelete) {{
    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    sc.exe delete $svc 2>&1 | Out-Null
    # WMI fallback
    try {{
        $svcObj = Get-WmiObject -Query "SELECT * FROM Win32_Service WHERE Name='$svc'" -ErrorAction SilentlyContinue
        if ($svcObj) {{
            $svcObj.Delete() | Out-Null
            Write-Log "Deleted service via WMI: $svc"
        }}
    }} catch {{
        Write-Log "WMI delete failed for $svc : $($_.Exception.Message)" "WARN"
    }}
    # Force-remove the service registry key directly
    $svcRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$svc"
    if (Test-Path $svcRegPath) {{
        Remove-Item -Path $svcRegPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Force-removed service registry key: $svcRegPath"
    }}
}}
Start-Sleep 3

# --- STEP 11: Delete SAAZOD and ITSPlatform folders ---
$foldersToDelete = @("C:\Program Files (x86)\SAAZOD", "C:\Program Files (x86)\ITSPlatform", "C:\Program Files (x86)\ITSPlatformSetupLogs")
foreach ($folder in $foldersToDelete) {{
    if (Test-Path $folder) {{
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Deleted folder: $folder"
    }}
}}

Write-Log "=== Cleanup Complete ==="

# =====================================================================
# PRE-INSTALL PREREQUISITE CHECKS
# Verify that all old services are truly gone before running installer.
# If services are stuck in "pending deletion" the MSI will fail with
# Error 1920 (service failed to start) and rollback (code 1603).
# =====================================================================
Write-Log "=== Running Pre-Install Prerequisite Checks ==="
$prereqFailed = $false

# --- CHECK 1: Verify old services are fully deleted ---
Write-Log "CHECK 1: Verifying old services are fully removed..."
$ghostServices = @()
$svcNamesToCheck = @("ITSPlatform", "ITSPlatformManager")
foreach ($svcName in $svcNamesToCheck) {{
    # Check via Get-Service
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($svc) {{
        Write-Log "  BLOCKED: Service '$svcName' still exists (Status: $($svc.Status))" "ERROR"
        $ghostServices += $svcName
    }}
    # Check via registry (service may be marked for deletion but key persists until reboot)
    $svcRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$svcName"
    if (Test-Path $svcRegPath) {{
        Write-Log "  BLOCKED: Service registry key still exists: $svcRegPath" "ERROR"
        $ghostServices += "$svcName (registry)"
    }}
}}

if ($ghostServices.Count -gt 0) {{
    Write-Log "PREREQUISITE FAILED: $($ghostServices.Count) ghost service(s) detected: $($ghostServices -join ', ')" "ERROR"
    Write-Log "These services are pending deletion and require a reboot to clear." "ERROR"
    Write-Log "ACTION REQUIRED: Reboot this machine and re-run the migration script." "ERROR"
    Write-Log "The scheduled task '$TaskName' has been kept so it can re-run after reboot." "WARN"
    Stop-Transcript
    exit 0
}}
Write-Log "  PASSED: No ghost services detected."

# --- CHECK 2: Verify no ITSPlatform processes are running ---
Write-Log "CHECK 2: Verifying no ITSPlatform processes are running..."
$blockingProcs = @()
$procsToCheck = @("platform-agent-core", "platform-agent-manager", "platform-communicator-tray", "platform-communicator-plugin", "platform-eventlog-plugin", "platform-sysevents-plugin")
foreach ($procName in $procsToCheck) {{
    $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($proc) {{
        Write-Log "  WARNING: Process '$procName' still running (PID: $($proc.Id)). Killing..." "WARN"
        Stop-Process -Name $procName -Force -ErrorAction SilentlyContinue
        $blockingProcs += $procName
    }}
}}
if ($blockingProcs.Count -gt 0) {{
    Write-Log "  Killed $($blockingProcs.Count) lingering process(es). Waiting 5s for cleanup..."
    Start-Sleep 5
}}
Write-Log "  PASSED: No blocking processes running."

# --- CHECK 3: Verify install directory is clean ---
Write-Log "CHECK 3: Verifying install directory is clean..."
if (Test-Path $InstallPath) {{
    $fileCount = (Get-ChildItem -Path $InstallPath -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-Log "  WARNING: Install path exists with $fileCount file(s). Removing..." "WARN"
    Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
    Start-Sleep 2
    if (Test-Path $InstallPath) {{
        Write-Log "  BLOCKED: Cannot delete install directory (files locked?)" "ERROR"
        $prereqFailed = $true
    }} else {{
        Write-Log "  Cleaned: Install directory removed successfully."
    }}
}} else {{
    Write-Log "  PASSED: Install directory is clean."
}}

# --- CHECK 4: Verify MSI file integrity ---
Write-Log "CHECK 4: Verifying MSI file integrity..."
if (-not (Test-Path $agentPath)) {{
    Write-Log "  BLOCKED: MSI file not found at: $agentPath" "ERROR"
    $prereqFailed = $true
}} else {{
    $msiSize = (Get-Item $agentPath).Length
    if ($msiSize -lt 1048576) {{  # Less than 1MB is suspicious for an MSI
        Write-Log "  WARNING: MSI file is only $msiSize bytes â€” possibly corrupt or incomplete download" "WARN"
    }} else {{
        Write-Log "  PASSED: MSI file exists ($msiSize bytes)."
    }}
}}

# --- CHECK 5: Verify no other MSI installation is in progress ---
Write-Log "CHECK 5: Checking for active MSI installations..."
$msiServerProc = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
if ($msiServerProc -and $msiServerProc.Count -gt 1) {{
    Write-Log "  WARNING: Multiple msiexec processes detected. Another install may be in progress." "WARN"
    Write-Log "  Waiting 60 seconds for other installer to finish..."
    Start-Sleep 60
    $msiServerProc = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
    if ($msiServerProc -and $msiServerProc.Count -gt 1) {{
        Write-Log "  BLOCKED: Other MSI install still running after 60s wait." "ERROR"
        $prereqFailed = $true
    }}
}} else {{
    Write-Log "  PASSED: No conflicting MSI installations."
}}

# --- CHECK 6: Verify sufficient disk space ---
Write-Log "CHECK 6: Checking disk space..."
$installDrive = (Split-Path $InstallPath -Qualifier)
$disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$installDrive'" -ErrorAction SilentlyContinue
if ($disk) {{
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    if ($freeGB -lt 1) {{
        Write-Log "  BLOCKED: Only $freeGB GB free on $installDrive â€” need at least 1 GB" "ERROR"
        $prereqFailed = $true
    }} else {{
        Write-Log "  PASSED: $freeGB GB free on $installDrive."
    }}
}} else {{
    Write-Log "  WARNING: Could not query disk space for $installDrive" "WARN"
}}

# --- CHECK 7: Verify running as SYSTEM or Admin ---
Write-Log "CHECK 7: Checking execution context..."
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {{
    Write-Log "  PASSED: Running as Administrator ($($currentUser.Name))."
}} else {{
    Write-Log "  BLOCKED: NOT running as Administrator. MSI install will fail." "ERROR"
    $prereqFailed = $true
}}

# --- PREREQUISITE GATE ---
if ($prereqFailed) {{
    Write-Log "=== ABORTING: One or more prerequisite checks FAILED ===" "ERROR"
    Write-Log "Review the log above and resolve blocking issues before re-running."
    Stop-Transcript
    exit 1
}}

Write-Log "=== All Prerequisite Checks PASSED ==="

# =====================================================================
# INSTALL WITH RETRY LOGIC
# If services don't fully provision within 5 minutes, re-run installer
# =====================================================================
$maxAttempts = 2
$attempt = 0
$installSuccess = $false

while ($attempt -lt $maxAttempts -and -not $installSuccess) {{
    $attempt++
    Write-Log "=== Install Attempt $attempt of $maxAttempts ==="

    $msiLog = Join-Path $WorkDir "msi_debug_attempt$attempt.log"
    $p = Start-Process msiexec.exe -ArgumentList "/i `"$agentPath`" $AgentArgs /lv* `"$msiLog`"" -Wait -PassThru

    if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {{
        Write-Log "MSI Install Successful (Code: $($p.ExitCode)) - Attempt $attempt"

        # =============================================================
        # WAIT 5 MINUTES FOR SERVICES TO FULLY BOOTSTRAP
        # =============================================================

        # --- Phase 1: Wait for ITSPlatform service to start (2 min) ---
        Write-Log "Waiting for ITSPlatform service to start..."
        $svcTimeout = 120
        $elapsed = 0
        $svcRunning = $false
        while ($elapsed -lt $svcTimeout) {{
            $svc = Get-Service -Name "ITSPlatform" -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -eq 'Running') {{
                Write-Log "ITSPlatform service is running (waited $elapsed seconds)."
                $svcRunning = $true
                break
            }}
            Start-Sleep 5
            $elapsed += 5
        }}
        if (-not $svcRunning) {{
            Write-Log "ITSPlatform service did not start within $svcTimeout seconds." "WARN"
        }}

        # --- Phase 2: Wait for ITSPlatformManager service (5 min) ---
        Write-Log "Waiting up to 5 minutes for ITSPlatformManager service to provision..."
        $mgrTimeout = 300
        $elapsed = 0
        $mgrRunning = $false
        while ($elapsed -lt $mgrTimeout) {{
            $mgr = Get-Service -Name "ITSPlatformManager" -ErrorAction SilentlyContinue
            if ($mgr -and $mgr.Status -eq 'Running') {{
                Write-Log "ITSPlatformManager service is running (waited $elapsed seconds)."
                $mgrRunning = $true
                break
            }}
            Start-Sleep 10
            $elapsed += 10
        }}

        # --- Evaluate: Are services fully provisioned? ---
        if ($svcRunning -and $mgrRunning) {{
            Write-Log "All services provisioned successfully on attempt $attempt."
            $installSuccess = $true
        }} else {{
            Write-Log "Services NOT fully provisioned after 5 minute wait (Attempt $attempt)." "WARN"
            if ($attempt -lt $maxAttempts) {{
                Write-Log "Stopping partial services and re-running installer..."
                Stop-Service -Name "ITSPlatform" -Force -ErrorAction SilentlyContinue
                Stop-Service -Name "ITSPlatformManager" -Force -ErrorAction SilentlyContinue
                Start-Sleep 10
            }}
        }}
    }} else {{
        Write-Log "MSI Install FAILED (Code: $($p.ExitCode)) - Attempt $attempt. Check $msiLog" "ERROR"
        if ($attempt -lt $maxAttempts) {{
            Write-Log "Will retry installer in 15 seconds..."
            Start-Sleep 15
        }}
    }}
}}

# --- Final Status Report ---
Write-Log "=== Final Service Status ==="
$allServices = Get-Service -Name "ITSPlatform*" -ErrorAction SilentlyContinue
if ($allServices) {{
    foreach ($s in $allServices) {{
        Write-Log "  $($s.Name) = $($s.Status)"
    }}
}} else {{
    Write-Log "  No ITSPlatform services found." "WARN"
}}

# --- Verify processes ---
$expectedProcs = @("platform-agent-core", "platform-agent-manager")
foreach ($procName in $expectedProcs) {{
    $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($proc) {{
        Write-Log "VERIFIED: Process '$procName' is running (PID: $($proc.Id))."
    }} else {{
        Write-Log "WARNING: Process '$procName' not found." "WARN"
    }}
}}

# --- Cleanup scheduled task only on success ---
if ($installSuccess -and (Test-Path $InstallPath)) {{
    Write-Log "Install path confirmed: $InstallPath"
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Scheduled task '$TaskName' removed."
    Write-Log "=== Migration Complete ==="
}} else {{
    Write-Log "=== Migration INCOMPLETE - services not fully provisioned after $maxAttempts attempts ===" "ERROR"
    Write-Log "Scheduled task kept for manual investigation."
}}

Stop-Transcript
'@

# Inject variables into template
$Runner = $RunnerTemplate -f $WorkDir, $LogFile, $Transcript, $NewAgentUrl, $NewAgentArgs, $NewAgentInstallPath, $UninstallExe, $TaskName

# Save script
$Runner | Set-Content $ScriptPath -Force -Encoding UTF8

# Detect OS version for compatibility
$osVersion = [System.Environment]::OSVersion.Version
$is2012R2 = ($osVersion.Major -eq 6 -and $osVersion.Minor -eq 3)

if ($is2012R2) {
    # =====================================================================
    # WINDOWS SERVER 2012 R2 COMPATIBLE SCHEDULED TASK REGISTRATION
    # =====================================================================
    $Action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

    $Trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date).AddMinutes(1)

    $Principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Principal $Principal `
        -Force
} else {
    # =====================================================================
    # WINDOWS SERVER 2016+ / WINDOWS 10+ SCHEDULED TASK REGISTRATION
    # =====================================================================
    $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
    $Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)
    $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force
}

Write-Host "Task registered. The downloaded file will keep its original filename from the server."
