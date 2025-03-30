<#
.SYNOPSIS
Disables power management and Wake-on-LAN settings via WMI and NetAdapter cmdlets for maximum performance.
#>

# Region: Helper functions for clear, aesthetic output
function Write-SectionHeader($Title) {
    # Writes a formatted section header to the console.
    # Cyan color is used for section titles for better visual separation.
    Write-Host "`n$(('-' * 5)) $Title $(('-' * (60 - $Title.Length)))" -ForegroundColor Cyan
}

function Write-Status($status, $message, $color = "White") {
    # Writes a status message with consistent padding and color.
    # $status: The status keyword (e.g., SUCCESS, ERROR, SKIPPED).
    # $message: The detailed message to display.
    # $color: The foreground color for the message (defaults to White).
    $padStatus = $status.PadRight(9) # Pad status for alignment
    Write-Host "[$padStatus] $message" -ForegroundColor $color
}

# Define consistent colors for statuses (VS Code Inspired Theme)
$colors = @{
    SUCCESS         = "Green"    # Success/Constants
    ERROR           = "Red"      # Errors
    SKIPPED         = "DarkGray" # Comments/Less important
    INFO            = "Cyan"     # Identifiers/General Info
    CONFIG          = "Yellow"   # Actions/Function calls
    FAILED          = "Red"      # Errors
    DISABLED        = "Green"    # Success/Constants
    NOT_APPLICABLE  = "DarkGray" # Comments/Less important
    WARNING         = "Yellow"   # Warnings
    DEV_ERROR       = "DarkRed"  # Distinct Error (avoiding Magenta)
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

# Helper function to check power setting values for both AC and DC (Improved Parsing & Return Value)
# Returns:
# $true if setting exists and value is correct
# $false if setting exists but value is incorrect (or parse failed)
# $null if setting does not exist
function Check-PowerSettingValue($Subgroup, $Setting, $ExpectedValue, $SettingDescription) {
    try {
        # Execute powercfg query, capturing all output streams (stdout and stderr)
        $powercfgOutput = & powercfg -query SCHEME_CURRENT $Subgroup $Setting *>&1 # Capture all streams

        # Check if query reported setting does not exist
        if ($LASTEXITCODE -ne 0 -and ($powercfgOutput -join ' ') -match 'setting specified does not exist') {
             # Log INFO only once here, not duplicating in the main loop
             Write-Status "INFO" "'$SettingDescription' setting ($Subgroup / $Setting) does not exist on this system." $colors.INFO
             return $null # Indicate setting does not exist
        }

        # Check for other non-zero exit codes or lack of expected output
        if ($LASTEXITCODE -ne 0 -and ($powercfgOutput -notmatch 'Current AC Power Setting Index')) {
             Write-Status "WARNING" ("Powercfg query failed for '{0}' ({1}/{2}). Output: {3}" -f $SettingDescription, $Subgroup, $Setting, ($powercfgOutput -join '; ')) $colors.WARNING
             return $false # Indicate failure to check properly, might need setting
        }

        # Process the output line by line to find the relevant indices, ignoring other lines
        $acValue = $null
        $dcValue = $null
        foreach ($line in $powercfgOutput) {
            if ($line -match 'Current AC Power Setting Index:\s+0x([0-9a-fA-F]+)') {
                $acValue = [int]$("0x" + $matches[1])
            } elseif ($line -match 'Current DC Power Setting Index:\s+0x([0-9a-fA-F]+)') {
                $dcValue = [int]$("0x" + $matches[1])
            }
        }

        # Check if both AC and DC values were found
        if ($acValue -ne $null -and $dcValue -ne $null) {
             # Values found, check if they match expected
             if ($acValue -eq $ExpectedValue -and $dcValue -eq $ExpectedValue) {
                return $true # Correct value
             } else {
                 return $false # Incorrect value
             }
        } else {
             # Log if we couldn't parse the values
             Write-Status "WARNING" ("Could not parse AC/DC values for '$SettingDescription' ({0}/{1}) from output: {2}" -f $SettingDescription, $Subgroup, $Setting, ($powercfgOutput -join '; ')) $colors.WARNING
             return $false # Indicate failure to check properly, might need setting
        }
    } catch {
        # Log if the command execution itself fails
        Write-Status "ERROR" ("Exception querying power setting {0} / {1}: {2}" -f $Subgroup, $Setting, $_.Exception.Message) $colors.ERROR
        return $false # Indicate failure to check properly, might need setting
    }
}

# Helper function to set power setting values for both AC and DC
function Set-PowerSettingValue($Subgroup, $Setting, $Value, $SettingDescription) {
    try {
        # Set AC value
        # Redirect stderr to null for set commands to suppress potential "Invalid Parameters" messages if they occur here too
        powercfg -SETACVALUEINDEX SCHEME_CURRENT $Subgroup $Setting $Value -ErrorAction Stop 2>$null
        if ($LASTEXITCODE -ne 0) { throw "powercfg -SETACVALUEINDEX returned non-zero exit code." }
        # Set DC value
        powercfg -SETDCVALUEINDEX SCHEME_CURRENT $Subgroup $Setting $Value -ErrorAction Stop 2>$null
         if ($LASTEXITCODE -ne 0) { throw "powercfg -SETDCVALUEINDEX returned non-zero exit code." }
        Write-Status "CONFIG" "Set '$SettingDescription' to $Value" $colors.CONFIG
        return $true
    } catch {
         # Use -f format operator for robust string construction
        Write-Status "ERROR" ("Failed to set '{0}' ({1} / {2}): {3}" -f $SettingDescription, $Subgroup, $Setting, $_.Exception.Message) $colors.ERROR
        return $false
    }
}

try {
    # Set High Performance mode
    $currentSchemeOutput = powercfg -getactivescheme
    # GUID for High Performance: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
    $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    if ($currentSchemeOutput -match $highPerfGuid) {
        Write-Status "SKIPPED" "High Performance power plan already active" $colors.SKIPPED
    } else {
        Write-Status "CONFIG" "Setting High Performance power plan" $colors.CONFIG
        powercfg -setactive $highPerfGuid
        # Verify change
        $newSchemeOutput = powercfg -getactivescheme
        if ($newSchemeOutput -match $highPerfGuid) {
            Write-Status "SUCCESS" "Successfully set High Performance power plan" $colors.SUCCESS
        } else {
            Write-Status "ERROR" "Failed to set High Performance power plan" $colors.ERROR
        }
    }

    # Define power settings to configure [Subgroup Alias/GUID, Setting Alias/GUID, Expected Value, Description]
    # Using GUIDs for reliability and language independence
    $powerSettings = @(
        # USB Selective Suspend Setting GUIDs
        @{Subgroup="2a737441-1930-4402-8d77-b2bebba308a3"; Setting="48e6b7a6-50f5-4782-a5d4-53bb8f07e226"; Value=0; Description="USB Selective Suspend"}
        # Link State Power Management GUIDs
        @{Subgroup="ee12f906-d277-404b-b6da-e5fa1a576df5"; Setting="501a4d13-42af-4429-9fd1-a8218c268e20"; Value=0; Description="PCIe Link State Power Management"}
        # Turn off hard disk after GUIDs
        @{Subgroup="0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting="6738e2c4-e8a5-4a42-b16a-e040e769756e"; Value=0; Description="Hard Disk Timeout (Minutes)"}
        # Minimum processor state GUIDs
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="893dee8e-2bef-41e0-89c6-b55d0929964c"; Value=100; Description="Minimum Processor State (%)"}
        # Processor Idle Promote Threshold GUIDs
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="468fe65e-e9d4-4dd0-b57c-f1aea7060ba8"; Value=0; Description="Processor Idle Promote Threshold"}
        # Processor Idle Demote Threshold GUIDs *** Skip Check due to parsing issues ***
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="7b224883-b3cc-4d79-819f-8374152cbe7c"; Value=0; Description="Processor Idle Demote Threshold"; SkipCheck=$true}
        # System cooling policy GUIDs (Active) *** Skip Check due to parsing issues ***
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="94d3a615-a899-4ac5-ae2b-e4d8f634367f"; Value=1; Description="System Cooling Policy (Active=1)"; SkipCheck=$true}
        # Maximum processor state GUIDs (Should already be 100% in High Perf, but ensure)
        @{Subgroup="54533251-82be-4824-96c1-47b60b740d00"; Setting="bc5038f7-23e0-4960-96da-33abaf5935ec"; Value=100; Description="Maximum Processor State (%)"}
        # Sleep after GUIDs (Disable Sleep)
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="29f6c1db-86da-48c5-9fdb-f2b67b1f44da"; Value=0; Description="Sleep After (Minutes)"}
        # Allow hybrid sleep GUIDs (Disable)
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="94ac6d29-73ce-41a6-809f-6363ba21b47e"; Value=0; Description="Hybrid Sleep"}
        # Allow wake timers GUIDs (Disable)
        @{Subgroup="238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting="bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d"; Value=0; Description="Wake Timers"}
        # Optional: Disable Display Timeout (Uncomment if needed)
        # @{Subgroup="7516b95f-f776-4464-8c53-06167f40cc99"; Setting="3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"; Value=0; Description="Display Off Timeout (Minutes)"}
    )

    # Settings that might not exist on all systems (Check existence before configuring)
    $optionalPowerSettings = @(
        # Network connectivity in Standby GUIDs (Modern Standby related)
        @{Subgroup="f15576e8-98b7-4186-b944-eafa664402d9"; Setting="8619b916-e004-4dd8-9b66-dae86f806698"; Value=0; Description="Network connectivity in Standby"} # Check under SUB_NONE if Energy Saver subgroup doesn't exist
        # GPU Preference Policy GUIDs
        @{Subgroup="5fb4938d-1ee8-4b0f-9a3c-5036b0ab5c6c"; Setting="dd848b3a-b055-48f8-8218-86915500de77"; Value=1; Description="GPU Preference (High Performance=1)"}
    )

    # Process standard power settings
    foreach ($settingInfo in $powerSettings) {
        if ($settingInfo.SkipCheck -eq $true) {
            # Check was explicitly skipped due to known issues
            Write-Status "INFO" "Skipping check for '$($settingInfo.Description)' due to known query issues. Set operation will NOT be attempted." $colors.INFO
            # Do not attempt to set the value as the check (and likely set) is unreliable
            continue # Move to the next setting
        }

        # Perform the check
        $checkResult = Check-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description

        if ($checkResult -eq $true) {
            # Setting exists and value is correct
            Write-Status "SKIPPED" "'$($settingInfo.Description)' already set to $($settingInfo.Value)" $colors.SKIPPED
        } elseif ($checkResult -eq $false) {
            # Setting exists but value is wrong, or check failed (but setting likely exists)
            # Attempt to set the value
            Set-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
        }
        # If $checkResult is $null, it means the setting doesn't exist (logged by Check function), so we do nothing.
    }

    # Process optional power settings (check if they exist first)
    foreach ($settingInfo in $optionalPowerSettings) {
        # Check if the setting exists by trying to query it
        $settingExists = $false
        try {
            & powercfg -query SCHEME_CURRENT $settingInfo.Subgroup $settingInfo.Setting -ErrorAction Stop *>&1 | Out-Null # Capture all output, discard
             if ($LASTEXITCODE -ne 0) { throw "Query failed" } # Check exit code explicitly
            $settingExists = $true
        } catch {
             # Attempt check under SUB_NONE for CONNECTIVITYINSTANDBY if primary subgroup failed
             if ($settingInfo.Setting -eq "8619b916-e004-4dd8-9b66-dae86f806698") {
                 try {
                     & powercfg -query SCHEME_CURRENT SUB_NONE $settingInfo.Setting -ErrorAction Stop *>&1 | Out-Null
                     if ($LASTEXITCODE -ne 0) { throw "Query failed under SUB_NONE" }
                     $settingInfo.Subgroup = "SUB_NONE" # Update subgroup if found here
                     $settingExists = $true
                 } catch {
                     # No need to log here, handled below if $settingExists is still false
                 }
             }
        }

        if ($settingExists) {
             # Setting exists, now check its value
            $checkResult = Check-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
            if ($checkResult -eq $true) {
                 Write-Status "SKIPPED" "'$($settingInfo.Description)' already set to $($settingInfo.Value)" $colors.SKIPPED
            } elseif ($checkResult -eq $false) {
                 Set-PowerSettingValue $settingInfo.Subgroup $settingInfo.Setting $settingInfo.Value $settingInfo.Description
            }
             # Do nothing if $checkResult is $null (shouldn't happen here as we checked existence)

        } else {
             # Setting does not exist
             Write-Status "INFO" "'$($settingInfo.Description)' setting ($($settingInfo.Subgroup)/$($settingInfo.Setting) or SUB_NONE) not found on this system." $colors.INFO
        }
    }

    # Disable Hibernation
    $hibernationEnabled = $false
    try {
        # Check registry key HKLM:\SYSTEM\CurrentControlSet\Control\Power\HibernateEnabled
        $hibernateRegValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "HibernateEnabled" -ErrorAction SilentlyContinue
        if ($hibernateRegValue -and $hibernateRegValue.HibernateEnabled -eq 1) {
            $hibernationEnabled = $true
        }
    } catch {
        Write-Status "WARNING" "Could not check hibernation status via registry." $colors.WARNING
        # Fallback to powercfg -h (less reliable for scripting)
        $hibernationStatusOutput = powercfg -h | Out-String
        if ($hibernationStatusOutput -notmatch "Hibernation has not been enabled") {
             $hibernationEnabled = $true # Assume enabled if output doesn't explicitly say disabled/not enabled
        }
    }

    if (-not $hibernationEnabled) {
        Write-Status "SKIPPED" "Hibernation already disabled" $colors.SKIPPED
    } else {
        Write-Status "CONFIG" "Disabling Hibernation" $colors.CONFIG
        try {
            powercfg -h off -ErrorAction Stop 2>$null # Suppress stderr
             if ($LASTEXITCODE -ne 0) { throw "powercfg -h off returned non-zero exit code." }
            Write-Status "DISABLED" "Hibernation" $colors.DISABLED
        } catch {
            Write-Status "ERROR" ("Failed to disable hibernation: {0}" -f $_.Exception.Message) $colors.ERROR
        }
    }

    # Apply changes by setting the current scheme active again
    Write-Status "CONFIG" "Applying power scheme changes" $colors.CONFIG
    powercfg -SETACTIVE SCHEME_CURRENT 2>$null # Suppress stderr

    Write-Status "SUCCESS" "System power settings optimization attempt completed." $colors.SUCCESS
}
catch {
    # Catch errors during the power settings block
    Write-Status "ERROR" ("Failed during power settings optimization: {0}" -f $_.Exception.Message) $colors.ERROR
}

# Region: Disable Device Power Management via WMI
Write-SectionHeader "Disabling Device Power Management (WMI)"

try {
    # Get Plug and Play devices that are currently present and OK
    $devices = Get-PnpDevice | Where-Object { $_.Status -eq "OK" -and $_.Present -eq $true }
    Write-Status "INFO" "Found $($devices.Count) present devices with status OK." $colors.INFO
}
catch {
    Write-Status "ERROR" ("Failed to retrieve devices using Get-PnpDevice: {0}" -f $_.Exception.Message) $colors.ERROR
    $devices = @() # Ensure $devices is an empty array if the command fails
}

# Counters for WMI results
$wmiResults = @{Success=0; AlreadyDisabled=0; Failed=0; NotApplicable=0; VerificationFailed=0}

foreach ($device in $devices) {
    $deviceName = $device.FriendlyName
    # Clean up the PNPDeviceID for WMI query (replace '\' with '\\')
    $deviceID = $device.PNPDeviceID -replace '\\', '\\'
    # Escape special regex characters in device ID for use in LIKE query (basic escaping)
    $instanceNamePattern = ($deviceID -replace '\[', '[[]' -replace '%', '[%]' -replace '_', '[_]')

    # Query WMI for power management capability
    $powerMgmt = $null
    try {
         # Use -Filter with escaped pattern
         $powerMgmt = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -Filter "InstanceName LIKE '%$($instanceNamePattern)%'" -ErrorAction Stop
    } catch {
        # Handle cases where the query fails (e.g., permissions, WMI issues, invalid pattern after escape)
         $wmiResults.NotApplicable++
         continue # Skip to the next device
    }


    if ($powerMgmt) {
        # Handle cases where multiple instances might be returned (should be rare with specific ID)
        if ($powerMgmt -is [array]) { $powerMgmt = $powerMgmt[0] }

        # Check if power management is currently enabled
        if ($powerMgmt.Enable -eq $false) {
            # Already disabled
            Write-Status "SKIPPED" "$deviceName (WMI Power Mgmt Already Disabled)" $colors.SKIPPED
            $wmiResults.AlreadyDisabled++
        }
        else {
            # Attempt to disable power management
            try {
                # Use Set-CimInstance to modify the property
                Set-CimInstance -InputObject $powerMgmt -Property @{Enable=$false} -ErrorAction Stop
                # Short pause to allow change to apply
                Start-Sleep -Milliseconds 150

                # Verify the change
                $verify = $null
                try {
                     $verify = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -Filter "InstanceName LIKE '%$($instanceNamePattern)%'" -ErrorAction Stop
                     if ($verify -is [array]) { $verify = $verify[0] }
                } catch {
                     # Verification query failed, log warning
                     Write-Status "WARNING" "$deviceName (WMI Power Mgmt Verification Query Failed: $($_.Exception.Message))" $colors.WARNING
                     $wmiResults.VerificationFailed++ # Count this as verification failure
                     continue # Skip further verification logic
                }


                if ($verify -and $verify.Enable -eq $false) {
                    # Successfully disabled
                    Write-Status "DISABLED" "$deviceName (WMI Power Mgmt)" $colors.DISABLED
                    $wmiResults.Success++
                } else {
                    # Verification failed (setting still shows enabled or verify object is null)
                    Write-Status "WARNING" "$deviceName (WMI Power Mgmt Verification Failed - Still shows enabled or verify failed)" $colors.WARNING
                    $wmiResults.VerificationFailed++
                }
            } catch {
                # Failed to set the property
                Write-Status "FAILED" ("{0} (WMI Power Mgmt Set Failed: {1})" -f $deviceName, $_.Exception.Message) $colors.FAILED
                $wmiResults.Failed++
            }
        }
    }
    else {
        # No power management instance found for this device via this WMI class
        $wmiResults.NotApplicable++
    }
}

# WMI Summary
Write-SectionHeader "WMI Summary"
Write-Host "Disabled: $($wmiResults.Success)" -ForegroundColor $colors.DISABLED
Write-Host "Already Disabled: $($wmiResults.AlreadyDisabled)" -ForegroundColor $colors.SKIPPED
Write-Host "Verification Failed: $($wmiResults.VerificationFailed)" -ForegroundColor $colors.WARNING # Using Warning color
Write-Host "Set Failed: $($wmiResults.Failed)" -ForegroundColor $colors.FAILED
Write-Host "Not Applicable/No WMI Control: $($wmiResults.NotApplicable)" -ForegroundColor $colors.NOT_APPLICABLE

# Region: Disable NIC Power Management & Wake Features
Write-SectionHeader "Disabling NIC Power Management & Wake Features"

try {
    # Get all network adapters, including physical and virtual
    $networkAdapters = Get-NetAdapter -IncludeHidden -ErrorAction Stop
    Write-Status "INFO" "Found $($networkAdapters.Count) network adapters (including hidden)." $colors.INFO
}
catch {
    Write-Status "ERROR" ("Get-NetAdapter failed: {0}" -f $_.Exception.Message) $colors.ERROR
    $networkAdapters = @() # Ensure $networkAdapters is an empty array
}

# Counters for NIC results
# Initialize with AlreadyDisabled key
$nicResults = @{Success=0; Failed=0; AlreadyDisabled=0; NotSupported=0; DeviceError=0} # Added DeviceError counter

foreach ($adapter in $networkAdapters) {
    $adapterName = $adapter.Name
    $adapterInterfaceDesc = $adapter.InterfaceDescription

    # Skip adapters that are not UP or disconnected (likely virtual or not in use)
    # Also skip known virtual adapters like Loopback
    if ($adapter.Status -ne 'Up' -or $adapterInterfaceDesc -like "*Loopback*") {
         # No need to log skipped adapters unless debugging
         # Write-Status "SKIPPED" "$adapterName ($adapterInterfaceDesc) - Status is $($adapter.Status) or is Loopback" $colors.SKIPPED
         continue
    }

    try {
        # Check current power management settings for the adapter
        $adapterPowerInfo = Get-NetAdapterPowerManagement -Name $adapterName -IncludeHidden -ErrorAction Stop

        # Check specific Wake-on-LAN settings
        $wolMagicPacketDisabled = $adapterPowerInfo.WakeOnMagicPacket -eq $false
        $wolPatternDisabled = $adapterPowerInfo.WakeOnPattern -eq $false

        # Check the setting "Allow the computer to turn off this device to save power"
        # This corresponds to the ArpOffload and NSOffload properties (disable them)
        $arpOffloadDisabled = $adapterPowerInfo.ArpOffload -eq $false
        $nsOffloadDisabled = $adapterPowerInfo.NSOffload -eq $false

        # Check if all relevant settings are already disabled
        if ($wolMagicPacketDisabled -and $wolPatternDisabled -and $arpOffloadDisabled -and $nsOffloadDisabled) {
            Write-Status "SKIPPED" "$adapterName ($adapterInterfaceDesc) - Power saving features already disabled" $colors.SKIPPED
            $nicResults.AlreadyDisabled++
        } else {
            # Disable the features
            Write-Status "CONFIG" "Disabling Power saving features for $adapterName ($adapterInterfaceDesc)" $colors.CONFIG
            # Use Set-NetAdapterPowerManagement for broader control
            Set-NetAdapterPowerManagement -Name $adapterName -IncludeHidden `
                -ArpOffload:$false `
                -NSOffload:$false `
                -WakeOnMagicPacket:$false `
                -WakeOnPattern:$false `
                -ErrorAction Stop

             # Verify (optional but good practice)
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
                 $nicResults.Failed++ # Count as failed if verification fails
             }
        }
    }
    catch { # Catch block handles all errors now, check specific messages inside
         $errorMessage = $_.Exception.Message
         # Check for the specific "device not functioning" error
         if ($errorMessage -like "*A device attached to the system is not functioning*") {
             Write-Status "DEV_ERROR" "$adapterName ($adapterInterfaceDesc) - Device/Driver Error: Not functioning" $colors.DEV_ERROR
             $nicResults.DeviceError++
         }
         # Handle cases where the adapter might not support power management settings
         elseif ($_.Exception.InnerException -and ($_.Exception.InnerException.Message -like "*The parameter is incorrect*" -or $_.Exception.InnerException.Message -like "*not supported by the network adapter*")) {
              Write-Status "INFO" "$adapterName ($adapterInterfaceDesc) - Does not support these power management settings." $colors.INFO
              $nicResults.NotSupported++
         } elseif ($errorMessage -like "*No matching MSFT_NetAdapterPowerManagementSettingData object found*") {
             Write-Status "INFO" "$adapterName ($adapterInterfaceDesc) - Power management settings object not found." $colors.INFO
             $nicResults.NotSupported++
         }
         else {
             # Log other errors during Get/Set operations using -f operator
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
Write-Host "Device/Driver Error: $($nicResults.DeviceError)" -ForegroundColor $colors.DEV_ERROR # Added device error count
Write-Host "Other Failures/Verify Failed: $($nicResults.Failed)" -ForegroundColor $colors.FAILED

# Region: Finalization
Write-SectionHeader "Script Completed"
Write-Status "INFO" "Power optimization script finished." $colors.INFO

# Pause script execution if run directly in console for user to see output
if ($Host.Name -eq 'ConsoleHost' -and -not $PSScriptRoot) {
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    # Wait for a key press
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
}
