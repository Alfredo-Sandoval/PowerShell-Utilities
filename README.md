# Windows Power Management Optimizer Script

## Purpose

This PowerShell script attempts to optimize Windows power settings for maximum performance by disabling various power-saving features. It targets settings related to:

* The active power plan (sets High Performance)
* USB selective suspend
* PCIe Link State Power Management
* Hard disk timeouts
* Processor idle states and C-states
* System cooling policy
* Sleep, Hybrid Sleep, Hibernation, and Wake Timers
* Optional settings like GPU preference and network connectivity in standby (if available)
* Device-specific power management via WMI (`MSPower_DeviceEnable`)
* Network Interface Card (NIC) power saving features (`ArpOffload`, `NSOffload`, `WakeOnMagicPacket`, `WakeOnPattern`) via `NetAdapter` cmdlets.

It modifies settings using `powercfg.exe`, WMI, and PowerShell's `NetAdapter` module.

## 🚨 WARNING & DISCLAIMER 🚨

* **Increased Power Consumption:** Disabling power management features **will increase power consumption** and potentially generate more heat. This is expected behavior when optimizing for performance over efficiency.
* **Battery Life:** This script may **negatively impact battery life** on laptops. It is generally recommended for desktop systems or laptops plugged into AC power where maximum performance is the priority.
* **System Variations:** Not all settings exist or are controllable on all hardware/Windows versions. The script attempts to detect and skip non-existent settings. **Pay close attention to `[INFO]`, `[WARNING]`, `[ERROR]`, and `[DEV_ERROR]` messages** in the script's output to understand what was applied and what issues were encountered on *your* specific system.
* **Use at Your Own Risk:** The author(s) are not responsible for any issues that may arise from using this script. **Run this script at your own risk.** Thoroughly test your system after running the script.

## Requirements

* Windows Operating System
* **Administrator Privileges** (Run PowerShell as Administrator)
* PowerShell (Developed and tested on PowerShell 7.5.0, likely compatible with recent versions)

## How to Use

1.  Download the `Disable-WindowsPowerManagement.ps1` file (or clone the repository).
2.  Open **PowerShell as Administrator**.
3.  Navigate to the directory where you saved the script using the `cd` command (e.g., `cd C:\Users\YourUser\Downloads`).
4.  If you haven't run scripts downloaded from the internet before, you may need to adjust your PowerShell execution policy. To allow the script to run just for this PowerShell session, you can use:
    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    ```
5.  Run the script:
    ```powershell
    .\Disable-WindowsPowerManagement.ps1
    ```
6.  Review the output carefully for any `WARNING`, `ERROR`, or `DEV_ERROR` messages.

## Understanding the Output

The script provides color-coded status updates for clarity:

* `[SUCCESS]` / `[DISABLED]`: (Green) Operation completed successfully.
* `[SKIPPED]`: (DarkGray) Operation was skipped because the setting was already configured correctly, the target was not applicable (e.g., disconnected NIC), or the check was deliberately bypassed.
* `[INFO]`: (Cyan) Informational message (e.g., setting doesn't exist on the system, skipping check due to known issues).
* `[CONFIG]`: (Yellow) A configuration change is being applied or attempted.
* `[WARNING]`: (Yellow) A non-critical issue occurred (e.g., WMI setting verification failed, `powercfg` query parsing issue). The script usually continues. These often indicate system/driver quirks.
* `[ERROR]`: (Red) A script or command error occurred (e.g., failed to set a value, command failed).
* `[DEV_ERROR]`: (DarkRed) An error likely related to the specific device or its driver (e.g., NIC reported "device not functioning"). This usually indicates a problem outside the script itself that prevents management of that device.

## Customization

You can change the output color theme by modifying the `$colors` hashtable near the beginning of the script file. The values should correspond to valid `ForegroundColor` options in PowerShell (e.g., "Green", "Red", "Cyan", "Yellow", "DarkGray").

## License

This project is licensed under the MIT License.

---

**MIT License**

Copyright (c) 2025 [Your Name or GitHub Username]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
# Windows Power Management Optimizer Script

## Purpose

This PowerShell script attempts to optimize Windows power settings for maximum performance by disabling various power-saving features. It targets settings related to:

* The active power plan (sets High Performance)
* USB selective suspend
* PCIe Link State Power Management
* Hard disk timeouts
* Processor idle states and C-states
* System cooling policy
* Sleep, Hybrid Sleep, Hibernation, and Wake Timers
* Optional settings like GPU preference and network connectivity in standby (if available)
* Device-specific power management via WMI (`MSPower_DeviceEnable`)
* Network Interface Card (NIC) power saving features (`ArpOffload`, `NSOffload`, `WakeOnMagicPacket`, `WakeOnPattern`) via `NetAdapter` cmdlets.

It modifies settings using `powercfg.exe`, WMI, and PowerShell's `NetAdapter` module.

## 🚨 WARNING & DISCLAIMER 🚨

* **Increased Power Consumption:** Disabling power management features **will increase power consumption** and potentially generate more heat. This is expected behavior when optimizing for performance over efficiency.
* **Battery Life:** This script may **negatively impact battery life** on laptops. It is generally recommended for desktop systems or laptops plugged into AC power where maximum performance is the priority.
* **System Variations:** Not all settings exist or are controllable on all hardware/Windows versions. The script attempts to detect and skip non-existent settings. **Pay close attention to `[INFO]`, `[WARNING]`, `[ERROR]`, and `[DEV_ERROR]` messages** in the script's output to understand what was applied and what issues were encountered on *your* specific system.
* **Use at Your Own Risk:** The author(s) are not responsible for any issues that may arise from using this script. **Run this script at your own risk.** Thoroughly test your system after running the script.

## Requirements

* Windows Operating System
* **Administrator Privileges** (Run PowerShell as Administrator)
* PowerShell (Developed and tested on PowerShell 7.5.0, likely compatible with recent versions)

## How to Use

1.  Download the `Disable-WindowsPowerManagement.ps1` file (or clone the repository).
2.  Open **PowerShell as Administrator**.
3.  Navigate to the directory where you saved the script using the `cd` command (e.g., `cd C:\Users\YourUser\Downloads`).
4.  If you haven't run scripts downloaded from the internet before, you may need to adjust your PowerShell execution policy. To allow the script to run just for this PowerShell session, you can use:
    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    ```
5.  Run the script:
    ```powershell
    .\Disable-WindowsPowerManagement.ps1
    ```
6.  Review the output carefully for any `WARNING`, `ERROR`, or `DEV_ERROR` messages.

## Understanding the Output

The script provides color-coded status updates for clarity:

* `[SUCCESS]` / `[DISABLED]`: (Green) Operation completed successfully.
* `[SKIPPED]`: (DarkGray) Operation was skipped because the setting was already configured correctly, the target was not applicable (e.g., disconnected NIC), or the check was deliberately bypassed.
* `[INFO]`: (Cyan) Informational message (e.g., setting doesn't exist on the system, skipping check due to known issues).
* `[CONFIG]`: (Yellow) A configuration change is being applied or attempted.
* `[WARNING]`: (Yellow) A non-critical issue occurred (e.g., WMI setting verification failed, `powercfg` query parsing issue). The script usually continues. These often indicate system/driver quirks.
* `[ERROR]`: (Red) A script or command error occurred (e.g., failed to set a value, command failed).
* `[DEV_ERROR]`: (DarkRed) An error likely related to the specific device or its driver (e.g., NIC reported "device not functioning"). This usually indicates a problem outside the script itself that prevents management of that device.

## Customization

You can change the output color theme by modifying the `$colors` hashtable near the beginning of the script file. The values should correspond to valid `ForegroundColor` options in PowerShell (e.g., "Green", "Red", "Cyan", "Yellow", "DarkGray").

## License

This project is licensed under the MIT License.

---

**MIT License**

Copyright (c) 2025 [Your Name or GitHub Username]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
