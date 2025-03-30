<#
.SYNOPSIS
Disables power management and Wake-on-LAN settings via WMI and NetAdapter cmdlets.
#>

# Region: Helper functions for clear, aesthetic output
function Write-SectionHeader($Title) {
    Write-Host "`n$(('-' * 5)) $Title $(('-' * (60 - $Title.Length)))" -ForegroundColor Cyan
}

function Write-Status($status, $message, $color = "White") {
    $padStatus = $status.PadRight(9)
    Write-Host "[$padStatus] $message" -ForegroundColor $color
}

# Region: Check Administrator Privileges
Write-SectionHeader "Privilege Check"
if (-not (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator"))) {
    Write-Status "ERROR" "Administrator privileges required. Exiting." "Red"
    exit 1
} else {
    Write-Status "SUCCESS" "Script running with Administrator privileges." "Green"
}

# Region: Power Plan Optimization
Write-SectionHeader "Optimizing Power Settings"

# Add this helper function at the beginning to check power settings:
function Check-PowerSettingValue($Subgroup, $Setting, $ExpectedValue) {
    $currentValue = & powercfg -query SCHEME_CURRENT $Subgroup $Setting
    if ($currentValue -match "Current AC Power Setting Index: $ExpectedValue" -and 
        $currentValue -match "Current DC Power Setting Index: $ExpectedValue") {
        return $true
    }
    return $false
}

try {
    # Set High Performance mode
    $currentScheme = powercfg -getactivescheme
    if ($currentScheme -match "High performance") {
        Write-Status "SKIPPED" "High Performance power mode already active" "DarkGray"
    } else {
        Write-Status "ENABLED" "High Performance power mode" "Green"
        powercfg -setactive SCHEME_MIN
    }
    
    # Disable USB Selective Suspend
    if (Check-PowerSettingValue "2a737441-1930-4402-8d77-b2bebba308a3" "48e6b7a6-50f5-4782-a5d4-53bb8f07e226" 0) {
        Write-Status "SKIPPED" "USB Selective Suspend already disabled" "DarkGray"
    } else {
        Write-Status "CONFIG" "Disabling USB Selective Suspend" "White"
        powercfg -SETACVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
    }
    
    # Disable PCIe Link State Management
    if (Check-PowerSettingValue "SUB_PCIEXPRESS" "ASPM" 0) {
        Write-Status "SKIPPED" "PCIe Link State Management already disabled" "DarkGray"
    } else {
        Write-Status "CONFIG" "Disabling PCIe Link State Management" "White"
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PCIEXPRESS ASPM 0
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PCIEXPRESS ASPM 0
    }
    
    # Disable Disk Idle Timeout
    if (Check-PowerSettingValue "SUB_DISK" "DISKIDLE" 0) {
        Write-Status "SKIPPED" "Disk Idle Timeout already disabled" "DarkGray"
    } else {
        Write-Status "CONFIG" "Disabling Disk Idle Timeout" "White"
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_DISK DISKIDLE 0
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_DISK DISKIDLE 0
    }
    
    # Set Minimum Processor State to 100%
    if (Check-PowerSettingValue "SUB_PROCESSOR" "PROCTHROTTLEMIN" 100) {
        Write-Status "SKIPPED" "Minimum Processor State already set to 100%" "DarkGray"
    } else {
        Write-Status "CONFIG" "Setting Minimum Processor State to 100%" "White"
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 100
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 100
    }
    
    # Disable Hibernation
    $hibernationStatus = powercfg -h | Select-String -Pattern "Hibernation is disabled" -SimpleMatch
    if ($hibernationStatus) {
        Write-Status "SKIPPED" "Hibernation already disabled" "DarkGray"
    } else {
        Write-Status "DISABLED" "Hibernation" "Green"
        powercfg -h off
    }
    
    # Optional: Disable Display Timeout
    # Write-Status "CONFIG" "Disabling Display Timeout" "White"
    # powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0
    # powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0
    
    # NEW: Disable processor C-states (deeper sleep states)
    Write-Status "CONFIG" "Disabling processor C-states (deeper sleep states)" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEPROMOTE 0
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEPROMOTE 0
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEDEMOTE 0
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEDEMOTE 0
    
    # NEW: Set processor cooling policy to active (performance over quiet)
    Write-Status "CONFIG" "Setting processor cooling policy to active (performance)" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR SYSCOOLPOL 1
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR SYSCOOLPOL 1
    
    # NEW: Maximize processor performance state
    Write-Status "CONFIG" "Maximizing processor performance state policies" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PERFINCPOL 2
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PERFINCPOL 2
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PERFDECPOL 1
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PERFDECPOL 1
    
    # NEW: Disable sleep timeouts
    Write-Status "CONFIG" "Disabling sleep timeouts" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0
    
    # NEW: Disable hybrid sleep
    Write-Status "CONFIG" "Disabling hybrid sleep" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP HYBRIDSLEEP 0
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP HYBRIDSLEEP 0
    
    # NEW: Disable wake timers
    Write-Status "CONFIG" "Disabling wake timers" "White"
    powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP RTCWAKE 0
    powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP RTCWAKE 0
    
    # NEW: Disable connected standby (modern standby) - Only try if supported
    Write-Status "CONFIG" "Disabling connected standby" "White"
    $connectedStandbyCheck = powercfg -query SCHEME_CURRENT | Select-String -Pattern "Energy Saver"
    if ($connectedStandbyCheck) {
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_ENERGYSAVER ESPOLICY 0
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_ENERGYSAVER ESPOLICY 0
    } else {

        Write-Status "SKIPPED" "Connected standby not supported on this system" "DarkGray"
    }
    
    # NEW: Set GPU preference to high performance - Only try if supported 
    Write-Status "CONFIG" "Setting GPU preference to high performance" "White"
    $gpuCheck = powercfg -query SCHEME_CURRENT | Select-String -Pattern "Graphics"
    if ($gpuCheck) {
        # Use 1 instead of 2 as value - based on the error message
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_GRAPHICS GPUPREFERENCEPOLICY 1 
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_GRAPHICS GPUPREFERENCEPOLICY 1
    } else {

        Write-Status "SKIPPED" "GPU preference settings not supported on this system" "DarkGray"
    }
    
    # NEW: Prevent automatic network connectivity in standby - Only try if supported
    Write-Status "CONFIG" "Disabling automatic network connectivity in standby" "White"
    $networkStandbyCheck = powercfg -query SCHEME_CURRENT | Select-String -Pattern "Network connectivity in Standby"
    if ($networkStandbyCheck) {
        powercfg -SETACVALUEINDEX SCHEME_CURRENT SUB_NONE CONNECTIVITYINSTANDBY 0
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT SUB_NONE CONNECTIVITYINSTANDBY 0
    } else {

        Write-Status "SKIPPED" "Network standby settings not supported on this system" "DarkGray"
    }
    
    # Apply changes
    Write-Status "CONFIG" "Applying power scheme changes" "White"
    powercfg -SETACTIVE SCHEME_CURRENT
    
    Write-Status "SUCCESS" "System power settings optimized for maximum performance" "Green"
}
catch {
    Write-Status "ERROR" "Failed to optimize power settings: $($_.Exception.Message)" "Red"
}

# Region: Disable Power Management via WMI
Write-SectionHeader "Disabling Device Power Management (WMI)"

try {
    $devices = Get-PnpDevice | Where-Object { $_.Status -eq "OK" }
    Write-Status "INFO" "Found $($devices.Count) devices with status OK." "White"
}
catch {
    Write-Status "ERROR" "Failed to retrieve devices: $($_.Exception.Message)" "Red"
    $devices = @()
}

# Counters
$wmiResults = @{Success=0; AlreadyDisabled=0; Failed=0; NotApplicable=0}

foreach ($device in $devices) {
    $deviceName = $device.FriendlyName
    $deviceID = $device.PNPDeviceID

    $wql = "SELECT * FROM MSPower_DeviceEnable WHERE InstanceName LIKE '%$([Regex]::Escape($deviceID))%'"
    $powerMgmt = Get-CimInstance -Namespace root\wmi -Query $wql -ErrorAction SilentlyContinue

    if ($powerMgmt) {
        if ($powerMgmt.Enable -eq $false) {
            Write-Status "SKIPPED" "$deviceName (Already Disabled)" "DarkGray"
            $wmiResults.AlreadyDisabled++
        }
        else {
            Set-CimInstance -InputObject $powerMgmt -Property @{Enable=$false} -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            $verify = Get-CimInstance -Namespace root\wmi -Query $wql -ErrorAction SilentlyContinue

            if ($verify -and $verify.Enable -eq $false) {
                Write-Status "DISABLED" "$deviceName" "Green"
                $wmiResults.Success++
            } else {
                Write-Status "FAILED" "$deviceName (Verification Failed)" "Yellow"
                $wmiResults.Failed++
            }
        }
    }
    else {
        $wmiResults.NotApplicable++
    }
}

# WMI Summary
Write-SectionHeader "WMI Summary"
Write-Host "Disabled: $($wmiResults.Success)" -ForegroundColor Green
Write-Host "Already Disabled: $($wmiResults.AlreadyDisabled)" -ForegroundColor DarkGray
Write-Host "Not Applicable: $($wmiResults.NotApplicable)" -ForegroundColor Gray
Write-Host "Failures: $($wmiResults.Failed)" -ForegroundColor Yellow

# Region: Disable NIC Power Management
Write-SectionHeader "Disabling NIC Wake-on-LAN Features"

try {
    $networkAdapters = Get-NetAdapter -ErrorAction Stop
    Write-Status "INFO" "Found $($networkAdapters.Count) network adapters." "White"
}
catch {
    Write-Status "ERROR" "Get-NetAdapter failed: $($_.Exception.Message)" "Red"
    $networkAdapters = @()
}

# Counters
$nicResults = @{Success=0; Failed=0}

foreach ($adapter in $networkAdapters) {
    $adapterName = $adapter.Name
    try {
        $adapterPowerInfo = Get-NetAdapterPowerManagement -Name $adapterName -ErrorAction Stop
        if ($adapterPowerInfo.WakeOnMagicPacket -eq "Disabled" -and $adapterPowerInfo.WakeOnPattern -eq "Disabled") {
            Write-Status "SKIPPED" "$adapterName Wake-on-LAN features already disabled" "DarkGray"
            $nicResults.AlreadyDisabled++  # Add this counter to the $nicResults hash table
        } else {
            Disable-NetAdapterPowerManagement -Name $adapterName -WakeOnPattern -WakeOnMagicPacket -Confirm:$false -ErrorAction Stop
            Write-Status "DISABLED" "$adapterName Wake-on-LAN features" "Green"
            $nicResults.Success++
        }
    }
    catch {
        Write-Status "FAILED" "$adapterName ($($_.Exception.Message))" "Red"
        $nicResults.Failed++
    }
}

# NIC Summary (update to include already disabled)
Write-SectionHeader "NIC Summary"
Write-Host "Disabled: $($nicResults.Success)" -ForegroundColor Green
Write-Host "Already Disabled: $($nicResults.AlreadyDisabled)" -ForegroundColor DarkGray
Write-Host "Failures: $($nicResults.Failed)" -ForegroundColor Red

# Region: Finalization
Write-SectionHeader "Script Completed"
Write-Status "INFO" "All tasks completed." "White"

# Pause if run directly
if ($Host.Name -eq 'ConsoleHost' -and -not $PSScriptRoot) {
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
}
