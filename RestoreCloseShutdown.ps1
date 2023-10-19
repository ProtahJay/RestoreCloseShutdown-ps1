# ==============================================================================
# SCRIPT TO RESTORE MINIMIZED WINDOWS AND GRACEFULLY CLOSE ALL VISIBLE APPLICATIONS
# ==============================================================================

# -----------------------------------------------------------------------------
# STEP 1: Add the required type to access the 'user32.dll' Windows API functions.
# -----------------------------------------------------------------------------
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        // Import the 'ShowWindow' function from 'user32.dll'.
        // This function is used to set the specified window's show state.
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@ 

# -----------------------------------------------------------------------------
# STEP 2: Define the constant for restoring a window.
# The constant value '9' corresponds to SW_RESTORE, which restores a minimized 
# (or maximized) window.
# -----------------------------------------------------------------------------
$SW_RESTORE = 9

# -----------------------------------------------------------------------------
# STEP 3: Gracefully close all instances of Elgato Wave Link app if they're running.
# -----------------------------------------------------------------------------
$waveLinkProcesses = Get-Process | Where-Object { $_.ProcessName -eq "WaveLink" -and $_.Id -ne $PID }

foreach ($process in $waveLinkProcesses) {
    try {
        # 3.1: Use the 'ShowWindow' function to restore the application window if it's minimized.
        [void][User32]::ShowWindow($process.MainWindowHandle, $SW_RESTORE)

        # 3.2: Pause for a short duration to ensure the window has time to restore.
        Start-Sleep -Milliseconds 500

        # 3.3: Attempt to close the Wave Link application gracefully.
        $process.CloseMainWindow()

        # 3.4: Pause again to allow the application to process the close request.
        Start-Sleep -Seconds 1
    } catch {
        # 3.5: Handle any exceptions that might occur during the process.
        Write-Output "Error handling WaveLink process: $_"
    }
}

# -----------------------------------------------------------------------------
# STEP 4: Fetch all running processes that have a MainWindowTitle.
# This filters out background processes or those without a GUI.
# Additionally, we ensure we don't target the current PowerShell process.
# -----------------------------------------------------------------------------
$runningProcesses = Get-Process | Where-Object { $_.MainWindowTitle -ne '' -and $_.Id -ne $PID }

# -----------------------------------------------------------------------------
# STEP 5: Iterate over each identified process.
# -----------------------------------------------------------------------------
foreach ($process in $runningProcesses) {
    try {
        # 5.1: Use the 'ShowWindow' function to restore the application window if it's minimized.
        [void][User32]::ShowWindow($process.MainWindowHandle, $SW_RESTORE)

        # 5.2: Pause for a short duration to ensure the window has time to restore.
        Start-Sleep -Milliseconds 500

        # 5.3: Attempt to close the application gracefully using its main window.
        $process.CloseMainWindow()

        # 5.4: Pause again to allow the application to process the close request.
        Start-Sleep -Seconds 1
    } catch {
        # 5.5: Handle any exceptions that might occur during the process.
        Write-Output "Error handling process $($process.Name): $_"
    }
}

# -----------------------------------------------------------------------------
# STEP 6: Give all applications some additional time to close gracefully.
# -----------------------------------------------------------------------------
Start-Sleep -Seconds 2

# -----------------------------------------------------------------------------
# STEP 7: Initiate the shutdown of the Windows computer.
# This will forcefully shut down the machine, ensuring all operations cease.
# -----------------------------------------------------------------------------
Stop-Computer -Force
