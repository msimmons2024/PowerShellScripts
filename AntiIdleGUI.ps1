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

        [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
        public static extern short GetAsyncKeyState(int virtualKeyCode);

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

Function Get-WorkDayMinutesRemaining
{
	#Editing the $EndofDay variable will allow for adjustment of the end of the workday.
	$currentDaylight = (Get-Date).IsDaylightSavingTime()
	$currentTimeZone = Get-TimeZone
	if ($currentTimeZone.SupportsDaylightSavingTime)
	{
		if ($currentDaylight)
		{
			[datetime]$EndofDay = '22:00:05'
		}
		else
		{
			[datetime]$EndofDay = '23:00:05'
		}
	}
	else
	{
		[datetime]$EndofDay = '22:00:05'
	}
	if ($currentTimeZone.BaseUtcOffset -ne '00:00:00')
	{
		$EndofDay = $EndofDay.ToLocalTime()
	}
	$time = [System.TimeSpan]::FromTicks($endOfDay.Ticks - (Get-Date).Ticks)
	return $time.TotalMinutes
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Display Management'
$form.Size = New-Object System.Drawing.Size(500, 200)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
#$form.MinimizeBox = $false
$form.MaximizeBox = $false

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(320, 120)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = 'Done'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Location = New-Object System.Drawing.Point(20, 120)
$checkBox.Checked = $false
$checkBox.Text = "Restart at end of day"
$checkBox.Size = New-Object System.Drawing.Size(175, 23)
$form.Controls.Add($checkBox)

$font = New-Object System.Drawing.Font("Ariel", 12)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(480, 200)
$label.Text = 'Please enter the information in the space below:'
$label.Font = $font
$form.Controls.Add($label)

$form.Topmost = $true
$timer = New-Object System.Windows.Forms.Timer

$timer.Interval = 1000 # 1 seconds

$Pos = [System.Windows.Forms.Cursor]::Position
$PosDelta = 1
$closeScript = $true
$maxTimeInSeconds = 300
$wshell = New-Object -ComObject wscript.shell

[Flags()] enum ExecutionState
{
	ES_AWAYMODE_REQUIRED = 0x00000040
	ES_CONTINUOUS = 0x80000000
	ES_DISPLAY_REQUIRED = 0x00000002
	ES_SYSTEM_REQUIRED = 0x00000001
}

$threadResult = ([ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS + [ThreadModule.ExecutionState]::ES_DISPLAY_REQUIRED))
if ($threadResult -eq "ES_CONTINUOUS")
{
	$threadResult = [ExecutionState]::ES_DISPLAY_REQUIRED + [ExecutionState]::ES_CONTINUOUS
}
function Enable-DataGridViewDoubleBuffer
{
	param ([Parameter(Mandatory = $true)]
		[System.Windows.Forms.Control]$grid,
		[switch]$Disable)
	
	$type = $grid.GetType();
	$propInfo = $type.GetProperty("DoubleBuffered", ('Instance', 'NonPublic'))
	$propInfo.SetValue($grid, $Disable -eq $false, $null)
}

$timer.Add_Tick({
		
		$lastInputTime = [ThreadModule.ThreadManagement]::GetIdleTime()
		$time = New-TimeSpan -Seconds ($lastInputTime / 1000)
		$remainingMinutes = [System.Math]::Round((Get-WorkDayMinutesRemaining), 0)
		$minutesInDay = 60 * 9
		if ($remainingMinutes -le 0)
		{
			$percentComplete = 100
		}
		else
		{
			$percentComplete = [Math]::Round((($minutesInDay - $remainingMinutes) * 100) / $minutesInDay, 0)
			if ($percentComplete -le 0) { $percentComplete = 0 }
		}
		if ($time.TotalSeconds -gt $maxTimeInSeconds)
		{
			Write-Host "$(Get-Date) Generating activity, Press any key to exit:"
			$outlookTitle = Get-Process -Name *outlook* | Select-Object mainwindowtitle
			$teamsTitle = Get-Process -Name *teams* | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object mainwindowtitle
			Write-Host "Outlook: $($outlookTitle.MainWindowTitle)"
			$appResult = $wshell.AppActivate("$($outlookTitle.MainWindowTitle)")
			Write-Host "Teams: $($teamsTitle.MainWindowTitle)"
			$appResult = $wshell.AppActivate("$($teamsTitle.MainWindowTitle)")
			$wshell.SendKeys("{F15}")
			[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point((($Pos.X) + $PosDelta), $Pos.Y)
			[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point((($Pos.X) - $PosDelta), $Pos.Y)
		}
		if ($checkBox.Checked)
		{
			if ($remainingMinutes -le 0)
			{
				Write-Host "Shutting down"
				Restart-Computer -Force
			}
		}
		$label.Text = "Thread State: $($threadResult)`nTime since last input: $($time.Duration())`nWork day ends in: $($remainingMinutes) minutes`nPercent Complete: $($percentComplete)%"
		
	})

[Action]$form.add_Shown({
		
		Enable-DataGridViewDoubleBuffer $label
		
		$label.Text = "Thread State: $($threadResult)"
		
	})

$timer.Start()
$form.ShowDialog() | Out-Null
$timer.Stop()

[ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS) | Out-Null