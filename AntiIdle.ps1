Add-Type -TypeDefinition @"
 using System;
 using System.Diagnostics;
 using System.Runtime.InteropServices;

namespace ThreadModule
{
    public struct LASTINPUTINFO 
    {
        public uint cbSize;
        public uint dwTime;
    }
    public enum ExecutionState : uint
    {
        ES_AWAYMODE_REQUIRED = 0x00000040,
        ES_CONTINUOUS = 0x80000000,
        ES_DISPLAY_REQUIRED = 0x00000002,
        ES_SYSTEM_REQUIRED = 0x00000001
        // Legacy flag, should not be used.
        // ES_USER_PRESENT = 0x00000004
    }

    public static class ThreadManagement
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);

        [DllImport("User32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);        

        [DllImport("Kernel32.dll")]
        private static extern uint GetLastError();

        public static uint GetIdleTime()
        {
            LASTINPUTINFO lastInPut = new LASTINPUTINFO();
            lastInPut.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(lastInPut);
            GetLastInputInfo(ref lastInPut);

            return ((uint)Environment.TickCount - lastInPut.dwTime);
        }
    /// <summary>
    /// Get the Last input time in milliseconds
    /// </summary>
    /// <returns></returns>
        public static long GetLastInputTime()
        {
            LASTINPUTINFO lastInPut = new LASTINPUTINFO();
            lastInPut.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(lastInPut);
            if (!GetLastInputInfo(ref lastInPut))
            {
                throw new Exception(GetLastError().ToString());
            }       
            return lastInPut.dwTime;
        }
    }
}
"@


Add-Type -AssemblyName System.Windows.Forms
for ($i = 0; $i -lt 9; $i++)
{
    Write-Host ""
}
Write-Host "Enabling anti idle measures! Thread state now set to: " -NoNewline
[ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS + [ThreadModule.ExecutionState]::ES_DISPLAY_REQUIRED)
Write-Host "Press any key to exit: "
$Pos = [System.Windows.Forms.Cursor]::Position
$PosDelta = 1
$closeScript = $true
$maxTimeInSeconds = 300
$wshell = New-Object -ComObject wscript.shell
do
{
    if ([Console]::KeyAvailable)
    {
        $key = [Console]::ReadKey();
        if ($key.key -eq 'F15')
        {
        }
        else
        {
            $closeScript = $false
        }
    }
    else
    {
        $lastInputTime = [ThreadModule.ThreadManagement]::GetIdleTime()
        $time = New-TimeSpan -Seconds ( $lastInputTime / 1000 )
        if ($time.TotalSeconds -gt $maxTimeInSeconds)
        {
            Write-Host "$(Get-Date) Generating activity, Press any key to exit:"
            $outlookTitle = Get-Process -Name *outlook* | select mainwindowtitle
            $teamsTitle = Get-Process -Name *teams* | where { $_.MainWindowTitle -ne '' } | select mainwindowtitle
            Write-Host "Outlook: $($outlookTitle.MainWindowTitle)"
            $appResult = $wshell.AppActivate("$($outlookTitle.MainWindowTitle)")
            Write-Host "Teams: $($teamsTitle.MainWindowTitle)"
            $appResult = $wshell.AppActivate("$($teamsTitle.MainWindowTitle)")
            $wshell.SendKeys("{F15}")
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point((($Pos.X) + $PosDelta) , $Pos.Y)
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point((($Pos.X) - $PosDelta) , $Pos.Y)
        }
        else
        {
            Write-Progress -Activity "Waiting to generate activity, press any key to exit. Time since last input $($time.Duration())" -Status $true -PercentComplete ($time.TotalSeconds/$maxTimeInSeconds*100) -SecondsRemaining ($maxTimeInSeconds-$time.TotalSeconds)
        }
        Start-Sleep -Seconds 1
    }


} while ($closeScript)

Write-Host "Disabling anti idle measures, and exiting!"

[ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS)

