$scriptName = "AntiIdle.ps1"

function ScriptName() { return $MyInvocation.ScriptName; }
if ($PSCommandPath -eq $null) { function GetPSCommandPath() { return $MyInvocation.PSCommandPath; } $PSCommandPath = GetPSCommandPath; }
if ([String]::IsNullOrEmpty($PSCommandPath))
{
    $PSCommandPath = "$([System.Environment]::CurrentDirectory)\$($scriptName)"
}

Function Write-Log
{
	
	Param (
		[string]$text
	)
	
	Write-Host "$(Get-Date) $($text)"
	
	if ($LogPath -ne $null)
	{
		"$(Get-Date) $($text)" | Out-File $LogPath -Append
	}
	
	
}

Function Verify-Directory
{
	Param (
		[string]$fileFullPath,
		[switch]$CreateIfMissing = $true
	)
	
	$dirPath = [System.IO.Path]::GetDirectoryName($fileFullPath)
	
	if (!(Test-Path $dirPath))
	{
		if ($CreateIfMissing)
		{
			New-Item -ItemType Directory $dirPath
			Write-Log "Creating missing directory: $($dirPath)"
			return $true
		}
		return $false
	}
	else
	{
		return "Directory $($dirPath) exists"
	}
}

Add-Type -TypeDefinition @"
 using System;
 using System.Diagnostics;
 using System.Runtime.InteropServices;

namespace ThreadModule
{
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


    }
}
"@


$LogPath = "$($PSCommandPath)_log\$(get-date -f "yyyy-MM-dd_hh-mm_tt").log"
$TranscriptPath = "$($PSCommandPath)_log\transcript_$(get-date -f "yyyy-MM-dd_hh-mm_tt").log"


Verify-Directory $LogPath
Verify-Directory $TranscriptPath

Start-Transcript $TranscriptPath

Write-Log "Enabling anti idle measures!"

[ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS + [ThreadModule.ExecutionState]::ES_DISPLAY_REQUIRED)

$result = Read-Host "Press any key to disable anti idle and exit the script!"

Write-Log "Disabling anti idle measures, and exiting!"

[ThreadModule.ThreadManagement]::SetThreadExecutionState([ThreadModule.ExecutionState]::ES_CONTINUOUS)

Stop-Transcript