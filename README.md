# Windows Power Management Optimizer Script

## Purpose

This PowerShell script optimizes Windows for maximum performance by disabling power-saving features including:

* High Performance power plan activation
* USB selective suspend
* PCIe Link State Power Management d
* Hard disk timeouts
* Processor idle states and C-states
* System cooling policy
* Sleep, Hybrid Sleep, Hibernation, and Wake Timers
* GPU preference and network connectivity in standby (when available)
* Device power management via WMI (`MSPower_DeviceEnable`)
* NIC power saving features (`ArpOffload`, `NSOffload`, `WakeOnMagicPacket`, `WakeOnPattern`)

Implementation uses `powercfg.exe`, WMI, and PowerShell's `NetAdapter` module.

## Warning & Disclaimer

* **Increased Power Consumption:** Disabling power management features **will increase power consumption** and potentially generate more heat. This is expected behavior when optimizing for performance over efficiency.
* **Battery Life:** This script may **negatively impact battery life** on laptops. It is generally recommended for desktop systems or laptops plugged into AC power where maximum performance is the priority.
* **System Variations:** Not all settings exist or are controllable on all hardware/Windows versions. The script attempts to detect and skip non-existent settings. **Pay close attention to `[INFO]`, `[WARNING]`, `[ERROR]`, and `[DEV_ERROR]` messages** in the script's output to understand what was applied and what issues were encountered on *your* specific system.
* **Use at Your Own Risk:** The author(s) are not responsible for any issues that may arise from using this script. **Run this script at your own risk.** Thoroughly test your system after running the script.

## Requirements

* Windows Operating System
* **Administrator Privileges** (Run PowerShell as Administrator)
* PowerShell (Developed and tested on PowerShell 7.5.0, likely compatible with recent versions)

## How to Use

1. Download the `Disable-WindowsPowerManagement.ps1` file (or clone the repository).
2. Open **PowerShell as Administrator**.
3. Navigate to the directory where you saved the script using the `cd` command (e.g., `cd C:\Users\YourUser\Downloads`).
4. If you haven't run scripts downloaded from the internet before, you may need to adjust your PowerShell execution policy. To allow the script to run just for this PowerShell session, you can use:

    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    ```

5. Run the script:

    ```powershell
    .\Disable-WindowsPowerManagement.ps1
    ```

6. Review the output carefully for any `WARNING`, `ERROR`, or `DEV_ERROR` messages.

## License

This project is licensed under the MIT License.

---

**MIT License**

Copyright (c) 2025 Alfredo Sandoval

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
