<#
.SYNOPSIS
Disables power management and Wake-on-LAN settings via WMI and NetAdapter cmdlets for maximum performance.
#>

# Define consistent colors for statuses (VS Code Inspired Theme)
$colors = @{
    SUCCESS         = "Green"    
    ERROR           = "Red"      
    SKIPPED         = "DarkGray" 
    INFO            = "Cyan"     
    CONFIG          = "Yellow"   
    FAILED          = "Red"      
    DISABLED        = "Green"    
    NOT_APPLICABLE  = "DarkGray" 
    WARNING         = "Yellow"   
    DEV_ERROR       = "DarkRed"  
}

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
    Write-Status "ERROR" "Administrator privileges required. Exiting." $colors.ERROR
    exit 1
} else {
    Write-Status "SUCCESS" "Script running with Administrator privileges." $colors.SUCCESS
}

# Region: Power Plan Optimization
Write-SectionHeader "Optimizing Power Settings"

function Check-PowerSettingValue($Subgroup, $Setting, $ExpectedValue, $SettingDescription) {
    try {
        $powercfgOutput = & powercfg -query SCHEME_CURRENT $Subgroup $Setting *>&1

        if ($LASTEXITCODE -ne 0 -and ($powercfgOutput -join ' ') -match 'setting specified does not exist') {
             Write-Status "INFO" "'$SettingDescription' setting ($Subgroup / $Setting) does not exist on this system." $colors.INFO
             return $null
        }

        if ($LASTEXITCODE -ne 0 -and ($powercfgOutput -notmatch 'Current AC Power Setting Index')) {
             Write-Status "WARNING" ("Powercfg query failed for '{0}' ({1}/{2}). Output: {3}" -f $SettingDescription, $Subgroup, $Setting, ($powercfgOutput -join '; ')) $colors.WARNING
             return $false
        }

        $acValue = $null
        $dcValue = $null
        foreach ($line in $powercfgOutput) {
            if ($line -match 'Current AC Power Setting Index:\s+0x([0-9a-fA-F]+)') {
                $acValue = [int]$("0x" + $matches[1])
            } elseif ($line -match 'Current DC Power Setting Index:\s+0x([0-9a-fA-F]+)') {
                $dcValue = [int]$("0x" + $matches[1])
            }
        }

        if ($acValue -ne $null -and $dcValue -ne $null) {
             if ($acValue -eq $ExpectedValue -and $dcValue -eq $ExpectedValue) {
                return $true
             } else {
                 return $false
             }
        } else {
             Write-Status "WARNING" ("Could not parse AC/DC values for '$SettingDescription' ({0}/{1}) from output: {2}" -f $SettingDescription, $Subgroup, $Setting, ($powercfgOutput -join '; ')) $colors.WARNING
             return $false
        }
    } catch {
        Write-Status "ERROR" ("Exception querying power setting {0} / {1}: {2}" -f $Subgroup, $Setting, $_.Exception.Message) $colors.ERROR
        return $false
    }
}

function Set-PowerSettingValue($Subgroup, $Setting, $Value, $SettingDescription) {
    try {
        powercfg -SETACVALUEINDEX SCHEME_CURRENT $Subgroup $Setting $Value -ErrorAction Stop 2>$null
        if ($LASTEXITCODE -ne 0) { throw "powercfg -SETACVALUEINDEX returned non-zero exit code." }
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT $Subgroup $Setting $Value -ErrorAction Stop 2>$null
         if ($LASTEXITCODE -ne 0) { throw "powercfg -SETDCVALUEINDEX returned non-zero exit code." }
        Write-Status "CONFIG" "Set '$SettingDescription' to $Value" $colors.CONFIG
        return $true
    } catch {
        Write-Status "ERROR" ("Failed to set '{0}' ({1} / {2}): {3}" -f $SettingDescription, $Subgroup, $Setting, $_.Exception.Message) $colors.ERROR
        return $false
    }
}

try {
    $currentSchemeOutput = powercfg -getactivescheme
    $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    if ($currentSchemeOutput -match $highPerfGuid) {
        Write-Status "SKIPPED" "High Performance power plan already active" $colors.SKIPPED
    } else {
        Write-Status "CONFIG" "Setting High Performance power plan" $colors.CONFIG
        powercfg -setactive $highPerfGuid
        $newSchemeOutput = powercfg -getactivescheme
        if ($newSchemeOutput -match $highPerfGuid) {
            Write-Status "SUCCESS" "Successfully set High Performance power plan" $colors.SUCCESS
        } else {
            Write-Status "ERROR" "Failed to set High Performance power plan" $colors.ERROR
        }
    }

    $powerSettings = @(
        @{Subgroup="2a737441-1930-4402-8d77-b2bebba308a3"; Setting="48e6b7a6-50f5-4782-a5d4-53bb8f07e226"; Value=0; Description="USB Selective Suspend"}
        @{Subgroup="ee12f906-d277-404b-b6da-e5fa1a576df5"; Setting="501a4d13-42af-4429-9fd1-a8218c268e20"; Value=0; Description="PCIe Link State Power Management"}
        @{Subgroup="0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting="6738e2c4-e8a5-4a42-b16a-e040e769756e"; Value=0; Description="Hard Disk Timeout (Minutes)"}
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="893dee8e-2bef-41e0-89c6-b55d0929964c"; Value=100; Description="Minimum Processor State (%)"}
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="468fe65e-e9d4-4dd0-b57c-f1aea7060ba8"; Value=0; Description="Processor Idle Promote Threshold"}
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="7b224883-b3cc-4d79-819f-8374152cbe7c"; Value=0; Description="Processor Idle Demote Threshold"; SkipCheck=$true}
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="94d3a615-a899-4ac5-ae2b-e4d8f634367f"; Value=1; Description="System Cooling Policy (Active=1)"; SkipCheck=$true}
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="bc5038f7-23e0-4960-96da-33abaf5935ec"; Value=100; Description="Maximum Processor State (%)"}
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="29f6c1db-86da-48c5-9fdb-f2b67b1f44da"; Value=0; Description="Sleep After (Minutes)"}
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="94ac6d29-73ce-41a6-809f-6363ba21b47e"; Value=0; Description="Hybrid Sleep"}
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d"; Value=0; Description="Wake Timers"}
    )

    $optionalPowerSettings = @(
        @{Subgroup="f15576e8-98b7-4186-b944-eafa664402d9"; Setting="8619b916-e004-4dd8-9b66-dae86f806698"; Value=0; Description="Network connectivity in Standby"}
        @{Subgroup="5fb4938d-1ee8-4b0f-9a3c-5036b0ab5c6c"; Setting="dd848b3a-b055-48f8-8218-86915500de77"; Value=1; Description="GPU Preference (High Performance=1)"}
    )

    foreach ($settingInfo in $powerSettings) {
        if ($settingInfo.SkipCheck -eq $true) {
            Write-Status "INFO" "Skipping check for '$($settingInfo.Description)' due to known query issues. Set operation will NOT be attempted." $colors.INFO
            continue
        }

        $checkResult = Check-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description

        if ($checkResult -eq $true) {
            Write-Status "SKIPPED" "'$($settingInfo.Description)' already set to $($settingInfo.Value)" $colors.SKIPPED
        } elseif ($checkResult -eq $false) {
            Set-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
        }
    }

    foreach ($settingInfo in $optionalPowerSettings) {
        $settingExists = $false
        try {
            & powercfg -query SCHEME_CURRENT $settingInfo.Subgroup $settingInfo.Setting -ErrorAction Stop *>&1 | Out-Null
             if ($LASTEXITCODE -ne 0) { throw "Query failed" }
            $settingExists = $true
        } catch {
             if ($settingInfo.Setting -eq "8619b916-e004-4dd8-9b66-dae86f806698") {
                 try {
                     & powercfg -query SCHEME_CURRENT SUB_NONE $settingInfo.Setting -ErrorAction Stop *>&1 | Out-Null
                     if ($LASTEXITCODE -ne 0) { throw "Query failed under SUB_NONE" }
                     $settingInfo.Subgroup = "SUB_NONE"
                     $settingExists = $true
                 } catch {
                 }
             }
        }

        if ($settingExists) {
            $checkResult = Check-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
            if ($checkResult -eq $true) {
                 Write-Status "SKIPPED" "'$($settingInfo.Description)' already set to $($settingInfo.Value)" $colors.SKIPPED
            } elseif ($checkResult -eq $false) {
                 Set-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
            }
        } else {
             Write-Status "INFO" "'$($settingInfo.Description)' setting ($($settingInfo.Subgroup)/$($settingInfo.Setting) or SUB_NONE) not found on this system." $colors.INFO
        }
    }

    $hibernationEnabled = $false
    try {
        $hibernateRegValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "HibernateEnabled" -ErrorAction SilentlyContinue
        if ($hibernateRegValue -and $hibernateRegValue.HibernateEnabled -eq 1) {
            $hibernationEnabled = $true
        }
    } catch {
        Write-Status "WARNING" "Could not check hibernation status via registry." $colors.WARNING
        $hibernationStatusOutput = powercfg -h | Out-String
        if ($hibernationStatusOutput -notmatch "Hibernation has not been enabled") {
             $hibernationEnabled = $true
        }
    }

    if (-not $hibernationEnabled) {
        Write-Status "SKIPPED" "Hibernation already disabled" $colors.SKIPPED
    } else {
        Write-Status "CONFIG" "Disabling Hibernation" $colors.CONFIG
        try {
            powercfg -h off -ErrorAction Stop 2>$null
             if ($LASTEXITCODE -ne 0) { throw "powercfg -h off returned non-zero exit code." }
            Write-Status "DISABLED" "Hibernation" $colors.DISABLED
        } catch {
            Write-Status "ERROR" ("Failed to disable hibernation: {0}" -f $_.Exception.Message) $colors.ERROR
        }
    }

    Write-Status "CONFIG" "Applying power scheme changes" $colors.CONFIG
    powercfg -SETACTIVE SCHEME_CURRENT 2>$null

    Write-Status "SUCCESS" "System power settings optimization attempt completed." $colors.SUCCESS
}
catch {
    Write-Status "ERROR" ("Failed during power settings optimization: {0}" -f $_.Exception.Message) $colors.ERROR
}

# Region: Disable Device Power Management via WMI
Write-SectionHeader "Disabling Device Power Management (WMI)"

try {
    $devices = Get-PnpDevice | Where-Object { $_.Status -eq "OK" -and $_.Present -eq $true }
    Write-Status "INFO" "Found $($devices.Count) present devices with status OK." $colors.INFO
}
catch {
    Write-Status "ERROR" ("Failed to retrieve devices using Get-PnpDevice: {0}" -f $_.Exception.Message) $colors.ERROR
    $devices = @()
}

$wmiResults = @{Success=0; AlreadyDisabled=0; Failed=0; NotApplicable=0; VerificationFailed=0}

foreach ($device in $devices) {
    $deviceName = $device.FriendlyName
    $deviceID = $device.PNPDeviceID -replace '\\', '\\'
    $instanceNamePattern = ($deviceID -replace '\[', '[[]' -replace '%', '[%]' -replace '_', '[_]')

    $powerMgmt = $null
    try {
         $powerMgmt = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -Filter "InstanceName LIKE '%$($instanceNamePattern)%'" -ErrorAction Stop
    } catch {
         $wmiResults.NotApplicable++
         continue
    }

    if ($powerMgmt) {
        if ($powerMgmt -is [array]) { $powerMgmt = $powerMgmt[0] }

        if ($powerMgmt.Enable -eq $false) {
            Write-Status "SKIPPED" "$deviceName (WMI Power Mgmt Already Disabled)" $colors.SKIPPED
            $wmiResults.AlreadyDisabled++
        }
        else {
            try {
                Set-CimInstance -InputObject $powerMgmt -Property @{Enable=$false} -ErrorAction Stop
                Start-Sleep -Milliseconds 150

                $verify = $null
                try {
                     $verify = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -Filter "InstanceName LIKE '%$($instanceNamePattern)%'" -ErrorAction Stop
                     if ($verify -is [array]) { $verify = $verify[0] }
                } catch {
                     Write-Status "WARNING" "$deviceName (WMI Power Mgmt Verification Query Failed: $($_.Exception.Message))" $colors.WARNING
                     $wmiResults.VerificationFailed++
                     continue
                }

                if ($verify -and $verify.Enable -eq $false) {
                    Write-Status "DISABLED" "$deviceName (WMI Power Mgmt)" $colors.DISABLED
                    $wmiResults.Success++
                } else {
                    Write-Status "WARNING" "$deviceName (WMI Power Mgmt Verification Failed - Still shows enabled or verify failed)" $colors.WARNING
                    $wmiResults.VerificationFailed++
                }
            } catch {
                Write-Status "FAILED" ("{0} (WMI Power Mgmt Set Failed: {1})" -f $deviceName, $_.Exception.Message) $colors.FAILED
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
Write-Host "Disabled: $($wmiResults.Success)" -ForegroundColor $colors.DISABLED
Write-Host "Already Disabled: $($wmiResults.AlreadyDisabled)" -ForegroundColor $colors.SKIPPED
Write-Host "Verification Failed: $($wmiResults.VerificationFailed)" -ForegroundColor $colors.WARNING
Write-Host "Set Failed: $($wmiResults.Failed)" -ForegroundColor $colors.FAILED
Write-Host "Not Applicable/No WMI Control: $($wmiResults.NotApplicable)" -ForegroundColor $colors.NOT_APPLICABLE

# Region: Disable NIC Power Management & Wake Features
Write-SectionHeader "Disabling NIC Power Management & Wake Features"

try {
    $networkAdapters = Get-NetAdapter -IncludeHidden -ErrorAction Stop
    Write-Status "INFO" "Found $($networkAdapters.Count) network adapters (including hidden)." $colors.INFO
}
catch {
    Write-Status "ERROR" ("Get-NetAdapter failed: {0}" -f $_.Exception.Message) $colors.ERROR
    $networkAdapters = @()
}

$nicResults = @{Success=0; Failed=0; AlreadyDisabled=0; NotSupported=0; DeviceError=0}

foreach ($adapter in $networkAdapters) {
    $adapterName = $adapter.Name
    $adapterInterfaceDesc = $adapter.InterfaceDescription

    if ($adapter.Status -ne 'Up' -or $adapterInterfaceDesc -like "*Loopback*") {
         continue
    }

    try {
        $adapterPowerInfo = Get-NetAdapterPowerManagement -Name $adapterName -IncludeHidden -ErrorAction Stop

        $wolMagicPacketDisabled = $adapterPowerInfo.WakeOnMagicPacket -eq $false
        $wolPatternDisabled = $adapterPowerInfo.WakeOnPattern -eq $false
        $arpOffloadDisabled = $adapterPowerInfo.ArpOffload -eq $false
        $nsOffloadDisabled = $adapterPowerInfo.NSOffload -eq $false

        if ($wolMagicPacketDisabled -and $wolPatternDisabled -and $arpOffloadDisabled -and $nsOffloadDisabled) {
            Write-Status "SKIPPED" "$adapterName ($adapterInterfaceDesc) - Power saving features already disabled" $colors.SKIPPED
            $nicResults.AlreadyDisabled++
        } else {
            Write-Status "CONFIG" "Disabling Power saving features for $adapterName ($adapterInterfaceDesc)" $colors.CONFIG
            Set-NetAdapterPowerManagement -Name $adapterName -IncludeHidden `
                -ArpOffload:$false `
                -NSOffload:$false `
                -WakeOnMagicPacket:$false `
                -WakeOnPattern:$false `
                -ErrorAction Stop

             $verifyPowerInfo = Get-NetAdapterPowerManagement -Name $adapterName -IncludeHidden -ErrorAction SilentlyContinue
             if ($verifyPowerInfo -and
                 $verifyPowerInfo.WakeOnMagicPacket -eq $false -and
                 $verifyPowerInfo.WakeOnPattern -eq $false -and
                 $verifyPowerInfo.ArpOffload -eq $false -and
                 $verifyPowerInfo.NSOffload -eq $false)
             {
                Write-Status "DISABLED" "$adapterName ($adapterInterfaceDesc) - Power saving features" $colors.DISABLED
                $nicResults.Success++
             } else {
                 Write-Status "WARNING" "$adapterName ($adapterInterfaceDesc) - Verification failed after attempting disable." $colors.WARNING
                 $nicResults.Failed++
             }
        }
    }
    catch {
         $errorMessage = $_.Exception.Message
         if ($errorMessage -like "*A device attached to the system is not functioning*") {
             Write-Status "DEV_ERROR" "$adapterName ($adapterInterfaceDesc) - Device/Driver Error: Not functioning" $colors.DEV_ERROR
             $nicResults.DeviceError++
         }
         elseif ($_.Exception.InnerException -and ($_.Exception.InnerException.Message -like "*The parameter is incorrect*" -or $_.Exception.InnerException.Message -like "*not supported by the network adapter*")) {
              Write-Status "INFO" "$adapterName ($adapterInterfaceDesc) - Does not support these power management settings." $colors.INFO
              $nicResults.NotSupported++
         } elseif ($errorMessage -like "*No matching MSFT_NetAdapterPowerManagementSettingData object found*") {
             Write-Status "INFO" "$adapterName ($adapterInterfaceDesc) - Power management settings object not found." $colors.INFO
             $nicResults.NotSupported++
         }
         else {
             Write-Status "FAILED" ("{0} ({1}) - Error: {2}" -f $adapterName, $adapterInterfaceDesc, $errorMessage) $colors.FAILED
             $nicResults.Failed++
         }
    }
}

# NIC Summary
Write-SectionHeader "NIC Summary"
Write-Host "Disabled: $($nicResults.Success)" -ForegroundColor $colors.DISABLED
Write-Host "Already Disabled: $($nicResults.AlreadyDisabled)" -ForegroundColor $colors.SKIPPED
Write-Host "Not Supported/Applicable: $($nicResults.NotSupported)" -ForegroundColor $colors.NOT_APPLICABLE
Write-Host "Device/Driver Error: $($nicResults.DeviceError)" -ForegroundColor $colors.DEV_ERROR
Write-Host "Other Failures/Verify Failed: $($nicResults.Failed)" -ForegroundColor $colors.FAILED

# Region: Finalization
Write-SectionHeader "Script Completed"
Write-Status "INFO" "Power optimization script finished." $colors.INFO

if ($Host.Name -eq 'ConsoleHost' -and -not $PSScriptRoot) {
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
}
