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
$NewAgentUrl  = "https://prod.setup.itsupport247.net/windows/BareboneAgent/32/Main-Palm_Beach_Opera_Windows_OS_ITSPlatform_TKN4a8260e7-fc13-4d04-908c-f0683e482fa6/MSI/setup"
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

# Force TLS 1.2 (required for Windows Server 2012 R2 where it is not the default)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- DOWNLOAD ---
try {{
    Write-Log "Downloading agent from: $AgentUrl"
    # Download to WorkDir and keep the original filename the server provides
    $response = Invoke-WebRequest -Uri $AgentUrl -UseBasicParsing -ErrorAction Stop

    # Get the original filename from Content-Disposition header
    $cd = $response.Headers['Content-Disposition']
    if ($cd -and $cd -match 'filename\s*=\s*"?([^";]+)"?') {{
        $originalFileName = [uri]::UnescapeDataString($Matches[1].Trim())
    }} elseif ($response.BaseResponse.ResponseUri) {{
        # Use the filename from the final redirected URL
        $originalFileName = [uri]::UnescapeDataString($response.BaseResponse.ResponseUri.Segments[-1].TrimEnd('/'))
    }} else {{
        # Last resort: use last segment of original URL
        $originalFileName = [uri]::UnescapeDataString(([uri]$AgentUrl).Segments[-1].TrimEnd('/'))
    }}

    # Ensure filename has .msi extension if not already present
    if ($originalFileName -and -not $originalFileName.EndsWith('.msi')) {{
        $originalFileName = $originalFileName + ".msi"
    }}

    $agentPath = Join-Path $WorkDir $originalFileName
    [IO.File]::WriteAllBytes($agentPath, $response.Content)
    Write-Log "File saved with original name: $originalFileName"
    Write-Log "Full path: $agentPath"
}} catch {{
    Write-Log "Download FAILED: $($_.Exception.Message)" "ERROR"
    exit 1
}}

# --- CLEANUP ---
Write-Log "=== Beginning Full Cleanup ==="

# --- STEP 1: Stop ITSPlatform service ---
Write-Log "Stopping ITSPlatform service..."
Stop-Service -Name "ITSPlatform" -Force -ErrorAction SilentlyContinue
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
$servicesToDelete = @("itsplatform", "itsplatformmanager", "itsplatform service", "itsplatformmanager service")
foreach ($svc in $servicesToDelete) {{
    # Try Stop + sc.exe delete
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

# --- INSTALL ---
Write-Log "Installing new agent..."
$msiLog = Join-Path $WorkDir "msi_debug.log"
# Triple-quotes around $agentPath to handle the comma safely
$p = Start-Process msiexec.exe -ArgumentList "/i `"$agentPath`" $AgentArgs /lv* `"$msiLog`"" -Wait -PassThru

if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {{
    Write-Log "Install Successful (Code: $($p.ExitCode))"

    # =====================================================================
    # WAIT FOR SERVICES TO FULLY BOOTSTRAP
    # The MSI installs the base agent, but ITSPlatform needs time to:
    #   1. Start the core service
    #   2. Phone home to the management server
    #   3. Download and install sub-components (plugins/services)
    # =====================================================================

    # --- Phase 1: Wait for the primary ITSPlatform service to start ---
    Write-Log "Waiting for ITSPlatform service to start..."
    $svcTimeout = 120   # 2 minutes max
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
        Write-Log "WARNING: ITSPlatform service did not start within $svcTimeout seconds." "WARN"
    }}

    # --- Phase 2: Wait for ITSPlatformManager service (child service) ---
    Write-Log "Waiting for ITSPlatformManager service to provision..."
    $mgrTimeout = 300   # 5 minutes max for sub-components
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
    if (-not $mgrRunning) {{
        Write-Log "WARNING: ITSPlatformManager service did not start within $mgrTimeout seconds." "WARN"
    }}

    # --- Phase 3: Verify expected processes are running ---
    Write-Log "Verifying agent processes..."
    Start-Sleep 15  # Brief extra settle time after services register
    $expectedProcs = @("platform-agent-core", "platform-agent-manager")
    foreach ($procName in $expectedProcs) {{
        $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($proc) {{
            Write-Log "VERIFIED: Process '$procName' is running (PID: $($proc.Id))."
        }} else {{
            Write-Log "WARNING: Process '$procName' not found." "WARN"
        }}
    }}

    # --- Final status report ---
    $allServices = Get-Service -Name "ITSPlatform*" -ErrorAction SilentlyContinue
    if ($allServices) {{
        Write-Log "=== Final Service Status ==="
        foreach ($s in $allServices) {{
            Write-Log "  $($s.Name) = $($s.Status)"
        }}
    }}

    if (Test-Path $InstallPath) {{
        Write-Log "Install path confirmed: $InstallPath"
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Scheduled task '$TaskName' removed."
    }}

    Write-Log "=== Migration Complete ==="
}} else {{
    Write-Log "Install FAILED (Code: $($p.ExitCode)). Check $msiLog" "ERROR"
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
    # Scheduled task action
    $Action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

    # Trigger
    $Trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date).AddMinutes(1)

    # Principal (2012 R2 compatible — requires LogonType ServiceAccount)
    $Principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    # Register task
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
