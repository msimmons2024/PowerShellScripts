<#	
	Usage - Start of powershell script that is calling it.
	Set-Variable "DefaultScriptLocation" -Value $MyInvocation.MyCommand.Path -Scope Global
	$scriptLocation = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
	."$($scriptLocation)\ScriptCore.ps1"
	$Core = [DefaultScript]::new()
	$global:Core = $Core
	$tempItem = [RequiredModuleItem]::new("ImportExcel")
	$Core.RequiredModuleItems.Add($tempItem)
	$Core.LoadRequiredModules()
	$Core.FinishScript()
#>

Function ConvertTo-GzipBase64String
{
	param (
		$text
	)
	
	$ms = New-Object System.IO.MemoryStream
	$cs = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
	$sw = New-Object System.IO.StreamWriter($cs)
	$sw.Write($text)
	$sw.Close();
	$s = [System.Convert]::ToBase64String($ms.ToArray())
	return $s
}

function ConvertFrom-GzipBase64String
{
	param (
		$text
	)
	$data = [System.Convert]::FromBase64String($text)
	$ms = New-Object System.IO.MemoryStream
	$ms.Write($data, 0, $data.Length)
	$ms.Seek(0, 0) | Out-Null
	$sr = New-Object System.IO.StreamReader(New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Decompress))
	$string = $sr.ReadToEnd()
	return $string
}
Function ConvertTo-KqlDataTableFromArray
{
	param (
		$psArray,
		$hasHeaders = $true
	)
	
	$outputString = ''
	
	$props = Get-Member -InputObject $psArray[0]
	
	$workWithProps = @{ }
	$ignoreProps = @{ }
	
	if ($props[0].TypeName -eq 'Microsoft.ActiveDirectory.Management.ADUser')
	{
		$ignoreProps.Add("Item", "Item")
		$ignoreProps.Add("WriteWarningStream", "WriteWarningStream")
		$ignoreProps.Add("WriteVerboseStream", "WriteVerboseStream")
		$ignoreProps.Add("WriteInformationStream", "WriteInformationStream")
		$ignoreProps.Add("WriteErrorStream", "WriteErrorStream")
		$ignoreProps.Add("WriteDebugStream", "WriteDebugStream")
	}
	
	foreach ($prop in $props)
	{
		if ($prop.TypeName -eq 'Microsoft.ActiveDirectory.Management.ADUser' -or $prop.TypeName -eq 'Selected.Microsoft.ActiveDirectory.Management.ADUser')
		{
			if ($ignoreProps.ContainsKey("$($prop.Name)"))
			{
				continue
			}
			
			if ($prop.Definition.Contains("ADPropertyValueCollection"))
			{
				if (-not $workWithProps.ContainsKey($prop.Name))
				{
					$workWithProps.Add($prop.Name, $prop.Name)
				}
				else
				{
					Write-Debug "Skipping Adding $($prop.Name)($($prop.MemberType)))"
				}
			}
			else
			{
				if ($prop.MemberType -eq 'Property')
				{
					if (-not $workWithProps.ContainsKey($prop.Name))
					{
						$workWithProps.Add($prop.Name, $prop.Name)
						Write-Debug "Adding $($prop.Name)($($prop.MemberType)))"
					}
					else
					{
						Write-Debug "Skipping Adding $($prop.Name)($($prop.MemberType)))"
					}
				}
				if ($prop.MemberType -eq 'NoteProperty')
				{
					if (-not $workWithProps.ContainsKey($prop.Name))
					{
						$workWithProps.Add($prop.Name, $prop.Name)
						Write-Debug "Adding $($prop.Name)($($prop.MemberType)))"
					}
					else
					{
						Write-Debug "Skipping Adding $($prop.Name)($($prop.MemberType)))"
					}
				}
			}
		}
		else
		{
			if ($ignoreProps.ContainsKey("$($prop.Name)"))
			{
				continue
			}
			
			if ($prop.MemberType -eq 'Property')
			{
				if (-not $workWithProps.ContainsKey($prop.Name))
				{
					$workWithProps.Add($prop.Name, $prop.Name)
					Write-Debug "Adding $($prop.Name)($($prop.MemberType)))"
				}
				else
				{
					Write-Debug "Skipping Adding $($prop.Name)($($prop.MemberType)))"
				}
			}
			elseif ($prop.MemberType -eq 'NoteProperty')
			{
				if (-not $workWithProps.ContainsKey($prop.Name))
				{
					$workWithProps.Add($prop.Name, $prop.Name)
					Write-Debug "Adding $($prop.Name)($($prop.MemberType)))"
				}
				else
				{
					Write-Debug "Skipping Adding $($prop.Name)($($prop.MemberType)))"
				}
			}
		}
	}
	
	$outputstring = "let psData = datatable("
	
	foreach ($prop in $workWithProps.Keys)
	{
		$outputString += "$($prop.Replace(' ', "_")):string,"
	}
	
	$outputString = $outputString.Substring(0, $outputString.Length - 1)
	$outputString += ") ["
	$outputString += "`n"
	
	foreach ($obj in $psArray)
	{
		foreach ($prop in $workWithProps.Keys)
		{
			$outputString += "`"$($obj.$prop)`","
		}
		$outputString += "`n"
	}
	
	$outputString += "];`n"
	
	return $outputString
}

Function ConvertTo-KqlDataTableFromExcelFile
{
param
    (
        [string]$InputFile,
        [switch]$SaveToClipboard = $true,
        [switch]$SaveToFile = $true,
        [string]$SaveToFileName = ''
    )

    $excelData = Import-Excel $InputFile

    $kqlTableString = KqlDataTableFromArray -psArray $excelData

    if($SaveToClipboard)
    {
        Set-Clipboard -Value $kqlTableString
    }
    if ($SaveToFile)
    {
        if ([string]::IsNullOrEmpty($SaveToFileName))
        {
            $SaveToFileName = "$([System.IO.Path]::GetTempFileName()).txt"
        }

        $kqlTableString | Out-File $SaveToFileName

        Write-Log "KQL table string saved to file: $($SaveToFileName)"

        Start-Process $SaveToFileName
    }
}

function Coalesce-ArrayColumns
{
	param
	(
		$arrayData
	)
	
	$customMembers = @{ }
	
	foreach ($item in $arrayData)
	{
		$members = $item | Get-Member -MemberType NoteProperty
		foreach ($member in $members)
		{
			if ($customMembers.ContainsKey($member.Name) -eq $false)
			{
				$customMembers.Add($member.Name, $member.Name)
			}
		}
	}
	
	foreach ($key in $customMembers.Keys)
	{
		foreach ($row in $arrayData)
		{
			if ((Get-Member -InputObject $row -Name $key) -eq $null)
			{
				$row | Add-Member -MemberType NoteProperty -Name $key -Value $null
			}
		}
	}
	
	return $arrayData
}

Function New-RateLimitWaitTime
{
	[cmdletbinding()]
	param (
		[System.TimeSpan]$MinimumDelay = [System.TimeSpan]::FromSeconds(1),
		[System.TimeSpan]$TimeFrame = [System.TimeSpan]::FromHours(1),
		[int]$RatePerTimeFrame = 1000
	)
	
	BEGIN { }
	
	PROCESS
	{
		$props = @{
			'MinimumDelay'  = $MinimumDelay;
			'TimeFrame'	    = $TimeFrame;
			'AvailableTime' = $TimeFrame.TotalMilliseconds;
			'DefaultCost'   = $TimeFrame.TotalMilliseconds / $RatePerTimeFrame;
			'Wait'		    = $MinimumDelay;
			'FirstUse'	    = [DateTime]::Now;
			'LastUse'	    = [DateTime]::Now;
			'RatePerTimeFrame' = $RatePerTimeFrame
		}
		
		$obj = New-Object -TypeName PSObject -Property $props
		$obj.PSObject.TypeNames.Insert(0, 'Custom.RateLimit')
		
	}
	END
	{
		Write-Output $obj
	}
}

Function Get-RateLimitWaitTime
{
	[cmdletbinding()]
	param (
		[parameter(mandatory = $true)]
		[PSObject]$RateLimit
	)
	
	BEGIN
	{
		
	}
	
	PROCESS
	{
		
		$TimeCost = $RateLimit.DefaultCost
		$RateLimit.Wait = $RateLimit.MinimumDelay
		
		$Compare = [System.TimeSpan]::FromTicks($RateLimit.FirstUse.AddMilliseconds($RateLimit.TimeFrame.TotalMilliseconds).Ticks - (get-date).Ticks)
		
		Write-Verbose "`$Compare = $($Compare)"
		
		if ($Compare.TotalMilliseconds -lt 0)
		{
			Write-Verbose "Time elasped setting AvailableTime to $($RateLimit.TimeFrame.TotalMilliseconds)"
			$RateLimit.AvailableTime = $RateLimit.TimeFrame.TotalMilliseconds
			$RateLimit.FirstUse = [DateTime]::Now
		}
		
		if (($RateLimit.AvailableTime - $TimeCost) -lt 0)
		{
			$RateLimit.Wait = [System.TimeSpan]::FromTicks($RateLimit.FirstUse.AddMilliseconds($RateLimit.TimeFrame.TotalMilliseconds).Ticks - (get-date).Ticks)
			$TimeCost = 0
		}
		
		$RateLimit.LastUse = [DateTime]::Now
		$RateLimit.AvailableTime -= $TimeCost
		
	}
	
	END
	{
		
		
		Write-Verbose "WaitTime = $($RateLimit.Wait), Using MS = $($TimeCost), MsToUseInQueue = $($RateLimit.AvailableTime), TimeFrame = $($RateLimit.TimeFrame), MaxRatePerTimeFrame = $($RateLimit.RatePerTimeFrame), TimeFrame LastUsed = $($RateLimit.LastUse)"
		Write-Output $RateLimit
		
	}
}

enum CmMessageType
{
	Info = 2
	Error = 1
	Debug = 3
	Verbose = 4
}

Function ConvertFrom-LogAnalyticsJson
{
	[CmdletBinding()]
	[OutputType([Object])]
	Param (
		[parameter(Mandatory = $true)]
		[string]$JSON
	)
	
	$data = ConvertFrom-Json $JSON
	$count = 0
	foreach ($table in $data.Tables)
	{
		$count += $table.Rows.Count
	}
	
	$objectView = New-Object object[] $count
	$i = 0;
	foreach ($table in $data.Tables)
	{
		foreach ($row in $table.Rows)
		{
			# Create a dictionary of properties
			$properties = @{ }
			for ($columnNum = 0; $columnNum -lt $table.Columns.Count; $columnNum++)
			{
				$properties[$table.Columns[$columnNum].name] = $row[$columnNum]
			}
			# Then create a PSObject from it. This seems to be *much* faster than using Add-Member
			$objectView[$i] = (New-Object PSObject -Property $properties)
			$null = $i++
		}
	}
	
	$objectView
}

function Test-Ipv4AddressInRange
{
	[cmdletbinding()]
	[outputtype([System.Boolean])]
	param (
		# IP Address to find.
		[parameter(Mandatory,
				   Position = 0)]
		[validatescript({
				([System.Net.IPAddress]$_).AddressFamily -eq 'InterNetwork'
			})]
		[string]$IPAddress,
		# Range in which to search using CIDR notation. (ippaddr/bits)
		[parameter(Mandatory,
				   Position = 1)]
		[validatescript({
				$IP = ($_ -split '/')[0]
				$Bits = ($_ -split '/')[1]
				(([System.Net.IPAddress]($IP)).AddressFamily -eq 'InterNetwork')
				
				if (-not ($Bits))
				{
					throw 'Missing CIDR notiation.'
				}
				elseif (-not (0 .. 32 -contains [int]$Bits))
				{
					throw 'Invalid CIDR notation. The valid bit range is 0 to 32.'
				}
			})]
		[alias('CIDR')]
		[string]$Range
	)
	
	# Split range into the address and the CIDR notation
	[String]$CIDRAddress = $Range.Split('/')[0]
	[int]$CIDRBits = $Range.Split('/')[1]
	
	# Address from range and the search address are converted to Int32 and the full mask is calculated from the CIDR notation.
	[int]$BaseAddress = [System.BitConverter]::ToInt32((([System.Net.IPAddress]::Parse($CIDRAddress)).GetAddressBytes()), 0)
	[int]$Address = [System.BitConverter]::ToInt32(([System.Net.IPAddress]::Parse($IPAddress).GetAddressBytes()), 0)
	[int]$Mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $CIDRBits))
	
	# Determine whether the address is in the range.
	if (($BaseAddress -band $Mask) -eq ($Address -band $Mask))
	{
		$true
	}
	else
	{
		$false
	}
}

function Test-FastPing
{
	param
	(
		# make parameter pipeline-aware
		[Parameter(Mandatory, ValueFromPipeline)]
		[string[]]$ComputerName,
		$TimeoutMillisec = 350
	)
	
	begin
	{
		# use this to collect computer names that were sent via pipeline
		[Collections.ArrayList]$bucket = @()
		
		# hash table with error code to text translation
		$StatusCode_ReturnValue =
		@{
			0 = 'Success'
			11001 = 'Buffer Too Small'
			11002 = 'Destination Net Unreachable'
			11003 = 'Destination Host Unreachable'
			11004 = 'Destination Protocol Unreachable'
			11005 = 'Destination Port Unreachable'
			11006 = 'No Resources'
			11007 = 'Bad Option'
			11008 = 'Hardware Error'
			11009 = 'Packet Too Big'
			11010 = 'Request Timed Out'
			11011 = 'Bad Request'
			11012 = 'Bad Route'
			11013 = 'TimeToLive Expired Transit'
			11014 = 'TimeToLive Expired Reassembly'
			11015 = 'Parameter Problem'
			11016 = 'Source Quench'
			11017 = 'Option Too Big'
			11018 = 'Bad Destination'
			11032 = 'Negotiating IPSEC'
			11050 = 'General Failure'
		}
		
		
		# hash table with calculated property that translates
		# numeric return value into friendly text
		
		$statusFriendlyText = @{
			# name of column
			Name	   = 'Status'
			# code to calculate content of column
			Expression = {
				# take status code and use it as index into
				# the hash table with friendly names
				# make sure the key is of same data type (int)
				$StatusCode_ReturnValue[([int]$_.StatusCode)]
			}
		}
		
		# calculated property that returns $true when status -eq 0
		$IsOnline = @{
			Name	   = 'Online'
			Expression = { $_.StatusCode -eq 0 }
		}
		
		# do DNS resolution when system responds to ping
		$DNSName = @{
			Name	   = 'DNSName'
			Expression = {
				if ($_.StatusCode -eq 0)
				{
					if ($_.Address -like '*.*.*.*')
					{ [Net.DNS]::GetHostByAddress($_.Address).HostName }
					else
					{ [Net.DNS]::GetHostByName($_.Address).HostName }
				}
			}
		}
	}
	
	process
	{
		# add each computer name to the bucket
		# we either receive a string array via parameter, or 
		# the process block runs multiple times when computer
		# names are piped
		$ComputerName | ForEach-Object {
			$null = $bucket.Add($_)
		}
	}
	
	end
	{
		# convert list of computers into a WMI query string
		$query = $bucket -join "' or Address='"
		
		Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" | Select-Object -Property Address, $IsOnline, $DNSName, $statusFriendlyText, *
	}
	
}

function Limit-ExcelColumnWidth
{
	param (
		$workbook
	)
	
	$count = 0
	foreach ($sheet in $workbook.Workbook.Worksheets)
	{
		for ($count = 1; $count -ne $sheet.Dimension.Columns; $count++)
		{
			$sheet.Column($count).Width = 25
		}
	}
	
	return $workbook
}

Function Write-Log
{
	
	Param (
		[string]$text,
		[bool]$save = $true,
		[CmMessageType]$type = [CmMessageType]::Info
	)

    if (Get-Variable 'Core' -ErrorAction Ignore)
    {
        $core.WriteCmEntry($type, $text, $save)
    }
    else
    {
        Write-Host $text
    }
	
}

Function Write-LogError
{
	
	Param (
		[string]$text,
		[bool]$save = $true,
		[CmMessageType]$type = [CmMessageType]::Error
	)
	$core.WriteCmEntry($type, $text, $save)
	
}

function Get-ExcelDateTime
{
	
	Param (
		[object]$dateTimeObject
	)
	
	if ($dateTimeObject -ne $null)
	{
		if ($dateTimeObject.GetType() -eq [datetime])
		{
			return $dateTimeObject.ToString("MM/dd/yyyy h:mm:ss tt")
		}
		elseif ($dateTimeObject.GetType() -eq [string])
		{
			if (![string]::IsNullOrEmpty($dateTimeObject))
			{
				try
				{
					$dt = [datetime]::Parse($dateTimeObject)
					return $dt.ToString("MM/dd/yyyy h:mm:ss tt")
				}
				catch
				{
					Write-Error -Message "$($_.Exception.Message) $($_.ScriptStackTrace)"
					return ""
				}
				
			}
			else
			{
				Write-Error -Message "Input object is unknown, returning empty $($_.ScriptStackTrace)"
				return ""
			}
		}
	}
	else
	{
		Write-Warning -Message "Input object is null $($_.ScriptStackTrace), returning empty"
		return ""
	}
}

Function Get-ArrayCount
{
	Param (
		[object]$dataObject
	)
	
	if ($dataObject -is [System.Array])
	{
		return $dataObject.Count
	}
	elseif (!$dataObject)
	{
		return 0
	}
	else
	{
		return 1
	}
}

class RequiredModuleItem {
	[string]$ModuleName
	[bool]$HasRequiredVersion
	[string]$RequiredVersion
	
	RequiredModuleItem()
	{
		$this.ModuleName = [string]::Empty
		$this.HasRequiredVersion = $false
		$this.RequiredVersion = [string]::Empty
	}
	
	RequiredModuleItem($moduleName)
	{
		$this.ModuleName = $moduleName
		$this.HasRequiredVersion = $false
		$this.RequiredVersion = [string]::Empty
	}
	
	RequiredModuleItem($moduleName, $requiredVersion)
	{
		$this.ModuleName = $moduleName
		$this.HasRequiredVersion = $true
		$this.RequiredVersion = $requiredVersion
	}
	
	Verify()
	{
		$Core = $global:Core
		
		$elevated = $Core.IsElevated()
		
		if ($elevated)
		{
			$scope = "AllUsers"
		}
		else
		{
			$scope = "CurrentUser"
		}
		
		if ($this.HasRequiredVersion)
		{
			$module = $null
			$module = Get-Module -ListAvailable -Name $this.ModuleName
			if ($Core.GetArrayCount($module) -gt 0)
			{
				$DoUpdate = $true
				foreach ($mod in $module)
				{
					if ($mod.Version -eq $this.RequiredVersion)
					{
						Write-Log "Found $($this.ModuleName) ($($this.RequiredVersion)) no update required."
						$DoUpdate = $false
						break
					}
				}
				if ($DoUpdate)
				{
					foreach ($mod in $module)
					{
						if ($mod.Version -ne $this.RequiredVersion)
						{
							Write-Log "Updating $($this.ModuleName) to $($this.RequiredVersion)"
							try
							{
								Update-Module -Name $mod.Name -RequiredVersion $this.RequiredVersion -Force -Confirm:$false -ErrorAction Stop
								if (!$?)
								{
									Write-Log "Error updating $($mod.Name) to $($this.RequiredVersion), attempting to install instead."
									Install-Module -Name $mod.Name -RequiredVersion $this.RequiredVersion -Force -Confirm:$false -Scope $scope -AllowClobber
								}
							}
							catch
							{
								Write-Log "Error updating $($mod.Name) to $($this.RequiredVersion), attempting to install instead."
								Install-Module -Name $mod.Name -RequiredVersion $this.RequiredVersion -Force -Confirm:$false -Scope $scope -AllowClobber
							}
						}
						else
						{
							Write-Log "No update needed for $($mod.Name)($($this.RequiredVersion))"
						}
					}
				}
			}
			else
			{
				Write-Log "Installing $($module.Name)($($this.RequiredVersion))"
				Install-Module -Name $module.Name -RequiredVersion $this.RequiredVersion -Force -Confirm:$false
			}
			Write-Log "Loading $($this.ModuleName) ($($this.RequiredVersion))."
			Import-Module -RequiredVersion $this.RequiredVersion -Name $this.ModuleName
		}
		else
		{
			$module = $null
			$module = Get-Module -ListAvailable -Name $this.ModuleName
			if ($Core.GetArrayCount($module) -gt 0)
			{
				$DoUpdate = $true
				foreach ($mod in $module)
				{
					$modLatest = Find-Module -Name $mod.Name
					if ($mod.Version -eq $modLatest.Version)
					{
						Write-Log "Found $($this.ModuleName) ($($mod.Version)) is equal to the latest of $($modLatest.Version) no update required."
						$DoUpdate = $false
						Write-Log "Loading $($this.ModuleName)."
						Import-Module -Name $this.ModuleName
						break
					}
				}
				if ($DoUpdate)
				{
					foreach ($mod in $module)
					{
						$Core.WriteCmEntry("PowerShell $($this.ModuleName) module is installed!, checking for updates")
						$modLatest = Find-Module -Name $this.ModuleName
						if ($mod.Version -ne $modLatest.Version)
						{
							Write-Log "Installing latest $($this.ModuleName) $($modLatest.Version) from $($this.ModuleName) ($($mod.Version))"
							Install-Module -Name $mod.Name -Force -Confirm:$false -Scope $scope -AllowClobber
							break;
						}
						else
						{
							Write-Log "No update needed for $($this.ModuleName)"
						}
					}
				}
			}
			else
			{
				Write-Log "Installing $($this.ModuleName)"
				Install-Module -Name $module.Name -Force -Confirm:$false -Scope $scope -AllowClobber
				Write-Log "Loading $($this.ModuleName)."
				Import-Module -Name $this.ModuleName
			}
		}
	}
}

class DefaultScript {
	
	[string]$ScriptName
	[string]$LogFile
	[string]$OutputFile
	[string]$TranscriptFile
	[bool]$ServerMode
	[string]$DriveLetter
	[string]$SystemType
	[system.diagnostics.stopwatch]$ScriptStopwatch
	[System.Collections.ArrayList]$RequiredModules = @()
	[System.Collections.Generic.List`1[RequiredModuleItem]]$RequiredModuleItems = @()
	[PSObject]$RateLimit
	[DateTime]$StartTime
	[int]$CoreVersion = 42
	[bool]$MultiThreadedMode = $false
	
	DefaultScript($scriptName,
		$logFile,
		$outputFile)
	{
		$this.ScriptName = $scriptName
		$this.LogFile = $logFile
		$this.OutputFile = $outputFile
		$this.ScriptStopwatch = [system.diagnostics.stopwatch]::StartNew()
		Start-Transcript $this.TranscriptFile
		$this.WriteCmEntry("Starting $($this.ScriptName).")
		$this.IsServerCheck()
	}
	
	DefaultScript($scriptName,
		$logFile,
		$outputFile,
		$transcriptFile)
	{
		$this.ScriptName = $scriptName
		$this.LogFile = $logFile
		$this.OutputFile = $outputFile
		$this.TranscriptFile = $transcriptFile
		$this.ScriptStopwatch = [system.diagnostics.stopwatch]::StartNew()
		Start-Transcript $this.TranscriptFile
		$this.WriteCmEntry("Starting $($this.ScriptName).")
		$this.IsServerCheck()
	}
	
	DefaultScript($scriptName,
		$outputFile)
	{
		$this.ScriptName = $scriptName
		$this.LogFile = "$($this.GetScriptDirectory())\logs\$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.TranscriptFile = "$($this.GetScriptDirectory())\logs\transcript_$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.OutputFile = $outputFile
		$this.ScriptStopwatch = [system.diagnostics.stopwatch]::StartNew()
		Start-Transcript $this.TranscriptFile
		$this.WriteCmEntry("Starting $($this.ScriptName).")
		$this.IsServerCheck()
	}
	
	DefaultScript($outputFile)
	{
		$this.ScriptName = $this.GetScriptName()
		$this.LogFile = "$($this.GetScriptDirectory())\logs\$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.TranscriptFile = "$($this.GetScriptDirectory())\logs\transcript_$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.OutputFile = $outputFile
		$this.ScriptStopwatch = [system.diagnostics.stopwatch]::StartNew()
		Start-Transcript $this.TranscriptFile
		$this.WriteCmEntry("Starting $($this.ScriptName).")
		$this.IsServerCheck()
	}
	
	DefaultScript()
	{
		$this.ScriptName = $this.GetScriptName()
		$this.LogFile = "$($this.GetScriptDirectory())\logs\$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.TranscriptFile = "$($this.GetScriptDirectory())\logs\transcript_$($this.ScriptName)\$(get-date -Format "yyyy-MM-dd_hh-mm_tt").log"
		$this.OutputFile = [string]::Empty
		$this.ScriptStopwatch = [system.diagnostics.stopwatch]::StartNew()
		Start-Transcript $this.TranscriptFile
		$this.WriteCmEntry("Starting $($this.ScriptName).")
		$this.IsServerCheck()
	}
	
	IsServerCheck()
	{
		$this.StartTime = (get-date)
		$osInfo = Get-WmiObject -Class Win32_OperatingSystem
		$this.DriveLetter = 'C'
		if ($osInfo.ProductType -eq '3')
		{
			$this.ServerMode = $true
			$this.SystemType = 'Server'
		}
		else
		{
			$this.ServerMode = $false
			$this.SystemType = 'Workstation'
		}
		
		$this.WriteCmEntry("Detected System Type: $($this.SystemType)")
	}
	
	[string]GetScriptLocation()
	{
		$scriptDir = "$((Get-Variable "DefaultScriptLocation" -Scope Global).Value)"
		if ([string]::IsNullOrEmpty($scriptDir))
		{
			return ($PWD)
		}
		else
		{
			return $scriptDir
		}
		
	}
	
	[string]GetScriptDirectory()
	{
		return [System.IO.Path]::GetDirectoryName("$($this.GetScriptLocation())")
	}
	
	[string]GetScriptName()
	{
		if ([string]::IsNullOrEmpty($this.ScriptName))
		{
			$this.ScriptName = [System.IO.Path]::GetFileNameWithoutExtension("$($this.GetScriptLocation())")
		}
		return $this.ScriptName
	}
	
	FinishScript()
	{
		$this.WriteCmEntry("Script total run time: $($this.ScriptStopwatch.Elapsed)")
		$this.ScriptStopwatch.Stop()
		Stop-Transcript
	}
	
	[string]GetWeekNumber([datetime]$DateTime = (Get-Date))
	{
		$calendar = [CultureInfo]::InvariantCulture.Calendar
		$dow = $calendar.GetDayOfWeek($DateTime)
		$tempObj = New-Object -TypeName PSObject
		$tempObj | Add-Member -MemberType NoteProperty -Name "FirstDateOfWeek" -Value $DateTime.AddDays(1 - ($DateTime.DayOfWeek.value__))
		$tempObj | Add-Member -MemberType NoteProperty -Name "LastDateOfWeek" -Value $DateTime.AddDays(1 - ($DateTime.DayOfWeek.value__)).AddDays(6)
		$tempObj | Add-Member -MemberType NoteProperty -Name "DayOfWeek" -Value $DateTime.DayOfWeek
		$tempObj | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $DateTime
		$tempObj | Add-Member -MemberType NoteProperty -Name "WeekNumber" -Value $calendar.GetWeekOfYear($DateTime, [Globalization.CalendarWeekRule]::FirstFullWeek, [DayOfWeek]::Monday)
		return $tempObj
	}
	
	WriteLog([string]$text)
	{
		$this.WriteLog($text, $true, $this.LogFile)
	}
	
	WriteLog([string]$text, [bool]$saveToFile)
	{
		$this.WriteLog($text, $saveToFile, $this.LogFile)
	}
	
	WriteLog([string]$text, [bool]$saveToFile, [string]$location)
	{
		Write-Host "$(Get-Date) $($text)"
		
		if ($saveToFile)
		{
			
			$saveLocation = $null
			if ([String]::IsNullOrEmpty($location))
			{
				$saveLocation = $this.LogFile
			}
			else
			{
				$saveLocation = $location
			}
			
			if ([string]::IsNullOrEmpty($saveLocation))
			{
				Write-Host "Error cannot write to log file, no log file path set." -ForegroundColor Red
				return
			}
			
			$dirResult = $this.VerifyDirectoryNoMessage($saveLocation)
			
			if ($dirResult)
			{
				"$(Get-Date) $($text)" | Out-File $saveLocation -Append
			}
		}
	}
	
	WriteCmFileOnly([string]$message)
	{
		$this.WriteCmEntry([CmMessageType]::Info, $message, $true, $this.LogFile, $false)
	}
	
	WriteCmEntry([string]$message)
	{
		$this.WriteCmEntry([CmMessageType]::Info, $message, $true, $this.LogFile, $true)
	}
	
	WriteCmEntryVerbose([string]$message)
	{
		$this.WriteCmEntry([CmMessageType]::Verbose, $message, $true, $this.LogFile, $true)
	}
	
	WriteCmEntryDebug([string]$message)
	{
		$this.WriteCmEntry([CmMessageType]::Debug, $message, $true, $this.LogFile, $true)
	}
	
	WriteCmEntry([CmMessageType]$type, [string]$message)
	{
		$this.WriteCmEntry($type, $message, $true, $this.LogFile, $true)
	}
	
	WriteCmEntry([string]$message, [bool]$saveToFile)
	{
		$this.WriteCmEntry([CmMessageType]::Info, $message, $true, $this.LogFile, $true)
	}
	
	WriteCmEntry([string]$message, [bool]$saveToFile, [bool]$displayMessage)
	{
		$this.WriteCmEntry([CmMessageType]::Info, $message, $true, $this.LogFile, $displayMessage)
	}
	
	WriteCmEntry([CmMessageType]$type, [string]$message, [bool]$saveToFile)
	{
		$this.WriteCmEntry($type, $message, $saveToFile, $this.LogFile, $true)
	}
	
	WriteCmEntry([CmMessageType]$type, [string]$message, [bool]$saveToFile, [bool]$displayMessage)
	{
		$this.WriteCmEntry($type, $message, $saveToFile, $this.LogFile, $displayMessage)
	}
	
	WriteCmEntry([CmMessageType]$type, [string]$message, [bool]$saveToFile, [string]$location, [bool]$displayMessage)
	{
		if ($displayMessage)
		{
			switch ($type.ToString())
			{
				"Debug" {
					Write-Debug "$(Get-Date) $($message)"
				}
				"Verbose"
				{
					Write-Verbose "$(Get-Date) $($message)"
				}
				default {
					Write-Host "$(Get-Date) $($message)"
				}
				
			}
		}
		
		if ($saveToFile)
		{
			$saveLocation = $null
			if ($saveToFile)
			{
				$saveLocation = $null
				if ([String]::IsNullOrEmpty($location))
				{
					$saveLocation = $this.LogFile
				}
				else
				{
					$saveLocation = $location
				}
				if ([string]::IsNullOrEmpty($saveLocation))
				{
					Write-Host "Error cannot write to log file, no log file path set." -ForegroundColor Red
					return
				}
			}
			$dirResult = $this.VerifyDirectoryNoMessage($saveLocation)
			
			if ($dirResult)
			{
				
				$mutex = New-Object 'Threading.Mutex' $false, "MyInterprocMutex"
				switch ($type.ToString())
				{
					"Error" {
						if ($this.MultiThreadedMode) {
							$mutex.waitone()
							"$((get-date).ToString("yyyyMMddThhmmss")) [ERROR]: $message" | Out-File $saveLocation -Append
							$mutex.ReleaseMutex()
						}
						else {
							"$((get-date).ToString("yyyyMMddThhmmss")) [ERROR]: $message" | Out-File $saveLocation -Append
						}
					}
					"Info" {
						if ($this.MultiThreadedMode) {
						$mutex.waitone()
						"$((get-date).ToString("yyyyMMddThhmmss")) [INFO]: $message" | Out-File $saveLocation -Append
						$mutex.ReleaseMutex()
						}
						else {
							"$((get-date).ToString("yyyyMMddThhmmss")) [INFO]: $message" | Out-File $saveLocation -Append
						}
					}
					"Debug" {
						if ($this.MultiThreadedMode) {
						$mutex.waitone()
						"$((get-date).ToString("yyyyMMddThhmmss")) [DEBUG]: $message" | Out-File $saveLocation -Append
						$mutex.ReleaseMutex()
						} else {
							"$((get-date).ToString("yyyyMMddThhmmss")) [DEBUG]: $message" | Out-File $saveLocation -Append
						}
					}
					"Verbose" {
						if ($this.MultiThreadedMode) {
						$mutex.waitone()
						"$((get-date).ToString("yyyyMMddThhmmss")) [VERBOSE]: $message" | Out-File $saveLocation -Append
						$mutex.ReleaseMutex()
						} else {
							"$((get-date).ToString("yyyyMMddThhmmss")) [VERBOSE]: $message" | Out-File $saveLocation -Append
						}
					}
				}
			}
			
		}
	}
	
	[string]ToExcelDateTime([object]$dateTimeObject)
	{
		if ($dateTimeObject -ne $null)
		{
			if ($dateTimeObject.GetType() -eq [datetime])
			{
				return $dateTimeObject.ToString("MM/dd/yyyy h:mm:ss tt")
			}
			elseif ($dateTimeObject.GetType() -eq [string])
			{
				if (![string]::IsNullOrEmpty($dateTimeObject))
				{
					try
					{
						$dt = [datetime]::Parse($dateTimeObject)
						return $dt.ToString("MM/dd/yyyy h:mm:ss tt")
					}
					catch
					{
						Write-Error -Message "$($_.Exception.Message) $($_.ScriptStackTrace)"
						return ""
					}
					
				}
				else
				{
					Write-Error -Message "Input object is unknown, returning empty $($_.ScriptStackTrace)"
					return ""
				}
			}
		}
		else
		{
			Write-Warning -Message "Input object is null $($_.ScriptStackTrace), returning empty"
			return ""
		}
		return ""
	}
	
	[int]GetArrayCount([object]$object)
	{
		if ($object -is [System.Array])
		{
			return $object.Count
		}
		elseif (!$object)
		{
			return 0
		}
		else
		{
			return 1
		}
	}
	
	[object]ReturnMatchingArrayItem($sourceArray,
		$fieldName,
		$fieldValue)
	{
		$result = $sourceArray | Where-Object { $_.$fieldName -eq $fieldValue }
		$return = New-Object -TypeName PSObject
		$return | Add-Member -MemberType NoteProperty -Name "Matched" -Value (($this.GetArrayCount($result)) -gt 0)
		
		if ($return.Matched)
		{
			$return | Add-Member -MemberType NoteProperty -Name "Data" -Value $result
		}
		
		if ($result -is [System.Array])
		{
			$return | Add-Member -MemberType NoteProperty -Name "IsArray" -Value $true
		}
		else
		{
			$return | Add-Member -MemberType NoteProperty -Name "IsArray" -Value $false
		}
		
		return $return
	}
	
	ExpandObjects($propertyPrefix, $rootObject, $currentObject)
	{
		$this.ExpandObjects($propertyPrefix, $rootObject, $currentObject, ",", $false)
	}
	
	ExpandObjects($propertyPrefix, $rootObject, $currentObject, $delimitator)
	{
		$this.ExpandObjects($propertyPrefix, $rootObject, $currentObject, $delimitator, $false)
	}
	
	ExpandObjects($propertyPrefix, $rootObject, $currentObject, $delimitator, [bool]$copyAll)
	{
		if ($null -ne $currentObject)
		{
			$Props = Get-Member -InputObject $currentObject -MemberType *Property
			
			
			foreach ($prop in $Props)
			{
				$currentName = $prop.Name
				if ($prop.Definition -like "System.Management.Automation.PSCustomObject*")
				{
					$this.ExpandObjects("$($propertyPrefix)_$($prop.Name)", $rootObject, $currentObject.$currentName, $delimitator, [bool]$copyAll)
				}
				elseif ($prop.Definition -match "Collection|\[\]")
				{
					$rootObject | Add-Member -Name "$($propertyPrefix)_$($prop.Name.trim())" -MemberType NoteProperty -Value "$($currentObject.$currentName -join "$($delimitator) ")" -Force
				}
				elseif ($prop.Definition -like "*datetime*")
				{
					$rootObject | Add-Member -Name "$($propertyPrefix)_$($prop.Name.trim())" -MemberType NoteProperty -Value $currentObject.$currentName
				}
				elseif ($prop.Definition -like "*string*")
				{
					$rootObject | Add-Member -Name "$($propertyPrefix)_$($prop.Name.trim())" -MemberType NoteProperty -Value "$($currentObject.$currentName)"
				}
				elseif ($CopyAll)
				{
					$rootObject | Add-Member -Name "$($propertyPrefix)_$($prop.Name.trim())" -MemberType NoteProperty -Value "$($currentObject.$currentName)" -Force
				}
				else
				{
					Write-Verbose "Unknown: $($prop.Name)=$($prop.Definition)"
					Write-Verbose "Unknown: $($prop.Name)=$($currentObject.$currentName)" -ErrorAction Continue -WarningAction Continue
				}
			}
		}
	}
	
	[string]Base64EncodeText([string]$text)
	{
		$Bytes = [System.Text.Encoding]::Unicode.GetBytes($text)
		$EncodedText = [Convert]::ToBase64String($Bytes)
		return $EncodedText
	}
	
	[string]Base64DecodeText([string]$text)
	{
		$DecodedText = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($text))
		return $DecodedText
	}
	
	[object]LoadDomainComputers([object]$domainDictionary,
		[string[]]$domainName)
	{
		return $this.LoadDomainComputers($domainDictionary, $domainName, "*")
	}
	
	[object]LoadDomainComputers([object]$domainDictionary,
		[string]$domainName, [string[]]$properties)
	{
		if (!$domainDictionary.ContainsKey($domainName))
		{
			Write-Log "Collecting all computers for domain $($domainName)"
			try
			{
				$tempAd = Get-ADComputer -Filter * -Properties $properties -Server $domainName -ErrorAction SilentlyContinue
				$tempAr = @{ }
				Write-Log "Processing all computers for domain $($domainName)"
				foreach ($ado in $tempAd)
				{
					$tempAr.add($ado.name, $ado)
				}
				if (!$domainDictionary.ContainsKey($domainName))
				{
					$domainDictionary.Add($domainName, $tempAr)
					return $domainDictionary
				}
			}
			catch
			{
				$tempAr = @{ }
				Write-Log "Failed to collect all computers for domain $($domainName)"
				if (!$domainDictionary.ContainsKey($domainName))
				{
					$domainDictionary.Add($domainName, $tempAr)
					return $domainDictionary
				}
				
			}
			if (!$?)
			{
				$tempAr = @{ }
				Write-Log "Failed to collect all computers for domain $($domainName)"
				if (!$domainDictionary.ContainsKey($domainName))
				{
					$domainDictionary.Add($domainName, $tempAr)
					return $domainDictionary
				}
			}
		}
		return $domainDictionary
	}
	
	[bool]VerifyDirectoryNoMessage($fileFullPath)
	{
		return $this.VerifyDirectory($fileFullPath, $true, $false)
	}
	
	[bool]VerifyDirectory($fileFullPath)
	{
		return $this.VerifyDirectory($fileFullPath, $true, $true)
	}
	
	[bool]IsElevated()
	{
		[Security.Principal.WindowsPrincipal]$user = [Security.Principal.WindowsIdentity]::GetCurrent();
		return $user.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator);
	}
	
	LoadRequiredModules()
	{
		if ($this.GetArrayCount($this.RequiredModules) -eq 0)
		{
			$this.WriteCmEntry("No required modules to load")
			return
		}
		$this.WriteCmEntry("The following modules are required for $($this.ScriptName): $($this.RequiredModules -join ', ')")
		$installScope = 'CurrentUser'
		If (!$this.IsElevated())
		{
			$installScope = 'AllUsers'
		}
		
		try
		{
			$nugetPackage = Get-PackageProvider -Name "NuGet"
			
			if (!$?)
			{
				$this.WriteCmEntry([CmMessageType]::Error, "NuGet Error Detected, script exiting")
				$this.WriteCmEntry([CmMessageType]::Error, "$($Error[0].Exception.Message)")
				$this.WriteCmEntry([CmMessageType]::Error, "$($Error[0].ScriptStackTrace)")
				exit
			}
			
			if ($nugetPackage.Version.ToString() -lt "2.8.5.201")
			{
				$this.WriteCmEntry("NuGet Package version is out of date, updating to 2.8.5.201")
				Install-PackageProvider -Name "NuGet" -MinimumVersion "2.8.5.201" -Confirm:$false -Force -Scope $installScope
			}
			else
			{
				$this.WriteCmFileOnly("NuGet Package version is current.")
			}
		}
		catch [System.Exception]
		{
			$this.WriteCmEntry([CmMessageType]::Error, "Script exiting, NuGet Package error: $($_.Message)")
			exit
		}
		
		foreach ($mod in $this.RequiredModules)
		{
			$modules = Get-Module -ListAvailable -Name $mod
			If ($this.GetArrayCount($modules) -gt 0)
			{
				foreach ($mod1 in $modules)
				{
					$this.WriteCmEntry("PowerShell $($mod1.Name) module is installed!, checking for updates")
					$modLatest = Find-Module -Name $mod1
					if ($modLatest.Version -ne $mod1.Version)
					{
						
						$this.WriteCmEntry("Detected update for PowerShell $($mod1.Name) module, removing current version!")
						Uninstall-Module -Name $modLatest.Name -AllVersions -Force -Confirm:$false
						$this.WriteCmEntry("Downloading and installing the latest version of $($mod1.Name).")
						Install-Module -Name $modLatest.Name -Force -Confirm:$false -AllowClobber -Scope $installScope
						
					}
					else
					{
						$this.WriteCmEntry("No update found for PowerShell $($mod1.Name) module.")
					}
				}
			}
			Else
			{
				$modules = Find-Module -Name $mod
				foreach ($mod1 in $modules)
				{
					Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
					Write-Host "PowerShell $($mod1.Name) module is not installed, Installing!"
					Install-Module -Name $mod1.Name -Force -Confirm:$false -AllowClobber -Scope $installScope
					Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted
				}
				
				
			}
		}
	}
	
	LoadRequiredModulesV2()
	{
		foreach ($module in $this.RequiredModuleItems)
		{
			$module.Verify()
		}
	}
	
	DeGZipFile($infile,
		$outfile)
	{
		if ([String]::IsNullOrEmpty($outfile))
		{
			$outfile = ($infile -replace '\.gz$', '')
		}
		$input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
		$output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
		$gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
		$buffer = New-Object byte[](1024)
		while ($true)
		{
			$read = $gzipstream.Read($buffer, 0, 1024)
			if ($read -le 0) { break }
			$output.Write($buffer, 0, $read)
		}
		$gzipStream.Close()
		$output.Close()
		$input.Close()
	}
	
	[bool]VerifyDirectory([string]$fileFullPath,
		[bool]$CreateIfMissing = $true,
		[bool]$DisplayMessage)
	{
		$dirPath = [System.IO.Path]::GetDirectoryName($fileFullPath)
		
		if (!(Test-Path $dirPath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue))
		{
			if ($CreateIfMissing)
			{
				New-Item -ItemType Directory $dirPath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue | Out-Null
				if ($DisplayMessage)
				{
					$this.WriteCmEntry('Info', "Creating missing directory: $($dirPath)", $true, $true)
				}
				return $true
			}
			return $false
		}
		else
		{
			if ($DisplayMessage)
			{
				$this.WriteCmEntry('Info', "Directory $($dirPath) exists.", $true, $true)
			}
			return $true
		}
		return $false
	}
}