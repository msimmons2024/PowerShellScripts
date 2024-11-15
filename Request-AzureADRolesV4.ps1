<#
.SYNOPSIS
This script requires the ScriptCore.ps1 file in the same directory.
You should adjust the $EndofDay variables in Get-WorkDayMinutesRemaining to handle the end of your workday. My end of day is 5pm Central Time (US) 
.DESCRIPTION
This script can request all or the specified AAD or Azure resource roles.
.PARAMETER UserPrincipleName
This is the UPN of the azure ad account such as name@domain.com
.PARAMETER RequestRoles
RequestRoles will request all AAD Roles unless you specify roles to ignore or if you specific which ones to request using -RolesToRequest.
.PARAMETER RequestResource
RequestResources will request all Azure Resource Roles, if you specific which ones to request, it will only request them. -ResourcesToRequest.
.PARAMETER WorkDay
The Workday parameter will have the script request roles from whatever the current time is to the end of the workday + 15 minutes. If you are using the manual ScriptStartDay, that will be the starting time used for each additional day but not the first day.
.PARAMETER ResourcesToRequest
ResourcesToRequest specifies which Azure resource roles to request. You should pass it along as an array. @('ResourceRole1', 'ResourceRole2', etc..). This will only work if you have also passed along the -RequestResources switch.
.PARAMETER RolesToRequest
RolesToRequest specifies which Azure AAD roles to request. You should pass it along as an array. @('AADRole1', 'AADRole2', etc..). This will only work if you have also passed along the -RequestRoles switch.
.PARAMETER RolesToExclude
When using the -RequestRoles switch by default it will request all Azure AAD roles that you have assigned to your account. This allows you to specify which ones to ignore. The format should be an array @('AADRole1', 'AADRole2', etc...).
.PARAMETER RoleDurationInMinutes
From the current time and or defined start time, request the role for the durations in minutes specified. If the time is less then 5 minutes it will be automatically increased. You cannot request a role with a duration of less than five minutes.
.PARAMETER Reason
The justification for your role request
.PARAMETER WorkWeek
Will request roles until the day of week reaches Saturday. if you start the script on a weekend it will automatically turn this switch off.
.PARAMETER AllowWeekends
When you specify a ScriptStartDay and ScriptEndDay and would like weekends to be included you want to also include this switch.
.PARAMETER ScriptStartDay
The date time you would like to start request roles at in the following format: '06/03/2024 08:50:00', you also need to include an ScriptEndDay.
.PARAMETER ScriptEndDay
The day you'd like the script to end requesting roles on in the following format: '06/20/2024 08:50:00'.
.PARAMETER UseWAM
If the the Web Authentication Manager should be used or not, default is off.
.PARAMETER UseLatest
Use the latest requested powershell modules or the known good versions at the time of the script being developed. 
If use latest is used it will always use the latest version and update to the latest version if any new module updates are released.
.EXAMPLE
Example to request your roles for the current workday using the latest version of the required powershell modules.
."C:\Scripts\Request-AzureADRolesV4.ps1" -UserPrincipleName 'username@domain.com' -RequestRoles -RequestResource -RolesToExclude @('Service Support Administrator') -Reason 'Review security alerts for the day.' -Workday -UseLatest
.EXAMPLE
If you'd like to just request your roles for 180 minutes, use -RoleDurationInMinutes instead of -Workday
."C:\Scripts\Request-AzureADRolesV4.ps1" -UserPrincipleName 'username@domain.com' -RequestRoles -RequestResource -RolesToExclude @('Service Support Administrator') -Reason 'Review security alerts for the day.' -RoleDurationInMinutes 180 -UseLatest
.EXAMPLE
If you'd like to just request your roles until friday of the current week for the entire workday, use -Workday and -Workweek instead of -Workday
."C:\Scripts\Request-AzureADRolesV4.ps1" -UserPrincipleName 'username@domain.com' -RequestRoles -RequestResource -RolesToExclude @('Service Support Administrator') -Reason 'Review security alerts for the day.' -Workday -Workweek -UseLatest
.EXAMPLE
if you'd like to request your roles for between a date range and for a specified time you can do the following, keep in mind that request starttime for each day will depend on the ScriptStartDay StartTime
"C:\Scripts\Request-AzureADRolesV4.ps1" -UserPrincipleName 'username@domain.com' -RequestRoles -RequestResource -RolesToExclude @('Service Support Administrator') -Reason 'Review security alerts for the day.' -RoleDurationInMinutes 180 -ScriptStartDay '06/03/2024 08:50:00' -ScriptEndDay '06/07/2024 08:55:00' -UseLatest
#>

[CmdletBinding(SupportsShouldProcess = $false)]
Param (
	[Parameter(Mandatory = $true)]
	[string]$UserPrincipleName = "",
	[switch]$RequestRoles = $false,
	[switch]$RequestResource = $false,
	[switch]$WorkDay = $false,
	[string[]]$RolesToRequest = $null,
	[string[]]$ResourcesToRequest = $null,
	[string[]]$RolesToExclude = @(),
	[int]$RoleDurationInMinutes = 0,
	[switch]$WhatIf = $false,
	[Parameter(Mandatory = $true)]
	[string]$Reason = "",
	[string]$ScriptStartDay = [System.DateTime]::Now.ToString('o'),
	[string]$ScriptEndDay = [System.DateTime]::Now.ToString('o'),
	[switch]$WorkWeek = $false,
	[switch]$AllowWeekends = $false,
	[switch]$UseWAM = $false,
	[switch]$UseLatest = $false
)

Set-Variable "DefaultScriptLocation" -Value $MyInvocation.MyCommand.Path -Scope Global
$scriptLocation = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
."$($scriptLocation)\ScriptCore.ps1"

$Core = [DefaultScript]::new()
$global:Core = $Core

class AzureRoleOverrideTime
{
	[string]$RoleName
	[int]$DurationInMinutes
}

Function Get-WorkDayMinutesRemaining
{
	param (
		$currentDate = (Get-Date)
	)
	$currentDaylight = (Get-Date).IsDaylightSavingTime()
	$currentTimeZone = Get-TimeZone
	if ($currentTimeZone.SupportsDaylightSavingTime)
	{
		if ($currentDaylight)
		{
			[datetime]$EndofDay = "$($currentDate.ToShortDateString()) 22:30:00"
		}
		else
		{
			[datetime]$EndofDay = "$($currentDate.ToShortDateString()) 23:30:00"
		}
	}
	else
	{
		[datetime]$EndofDay = '22:30:00'
	}
	if ($currentTimeZone.BaseUtcOffset -ne '00:00:00')
	{
		$EndofDay = $EndofDay.ToLocalTime()
	}
	$time = [System.TimeSpan]::FromTicks($endOfDay.Ticks - $currentDate.Ticks)
	return [Math]::Round($time.TotalMinutes, 0)
}

function RequestAzureRole
{
	[CmdletBinding(SupportsShouldProcess = $false)]
	param (
		$roleId = "",
		$roleDisplayName = "",
		$roleResourceId = "",
		$roleProvider = "",
		[int]$addStartTimeMinutes = 0,
		[int]$DurationInMinutes = 0,
		[int]$MaxTimeAllowed,
		[string]$UserAdObjectId,
		[string]$UserPrincipleName,
		[switch]$useWorkDay = $false,
		[string]$requestReason,
		[int]$maxAttempts = 3,
		[int]$delayBetweenAttempts = 5,
		[System.DateTime]$endTime = [DateTime]::Now.AddMinutes(-30),
		[System.DateTime]$startDay = [DateTime]::Now
	)
	$Core.WriteCmEntry("Creating $($roleProvider) role request for $($roleDisplayName) and adding the following minutes $($addStartTimeMinutes) to the start time of $($startDay.ToString('o')). Workday status: $($WorkDay). The role allows a maximum of $($MaxTimeAllowed) minutes per request.")
	$dt = $startDay.ToUniversalTime().AddMinutes($addStartTimeMinutes)
	$endTime = $endTime.ToUniversalTime()
	
	if ($endTime.Ticks -ge $dt.Ticks)
	{
		$dt = $endTime.AddSeconds(1)
	}
	
	$startTime = $dt
	
	$schedule = New-Object Microsoft.Open.MSGraph.Model.AzureADMSPrivilegedSchedule
	$schedule.Type = "Once"
	$schedule.StartDateTime = $dt.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
	if ($useWorkDay)
	{
		$workDayMinutes = Get-WorkDayMinutesRemaining -currentDate $startDay
		$DurationInMinutes = ($workDayMinutes - $addStartTimeMinutes)
		$Core.WriteCmEntry("Workday: $($Workday), Remaining minutes in workday: $($DurationInMinutes) - End of day $($startDay.AddMinutes($workDayMinutes).ToString())")
	}
	
	
	if ($DurationInMinutes -eq 0)
	{
		$minutesToAdd = $MaxTimeAllowed
	}
	else
	{
		if ($DurationInMinutes -gt $MaxTimeAllowed)
		{
			$minutesToAdd = $MaxTimeAllowed
		}
		else
		{
			$minutesToAdd = $DurationInMinutes
		}
	}
	
	#Azure role request must be at least 5 (Adjusted to 30) minutes.
	if ($minutesToAdd -gt 0 -and $minutesToAdd -le 30)
	{
		$minutesToAdd = 30
	}
	
	$schedule.endDateTime = $dt.AddMinutes($minutesToAdd).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
	$endTime = $dt.AddMinutes($minutesToAdd)
	$endTimeString = $endTime.ToLocalTime().ToString("o")
	$timeSpan = [TimeSpan]::FromTicks($endTime.Ticks - $dt.Ticks)
	$duration = [System.Xml.XmlConvert]::ToString($timeSpan)
	
	if ($WhatIf)
	{
		$Core.WriteCmEntry("WhatIf: Attemp ($($currentAttemp)/$($maxAttempts)) to request $($roleProvider) role $($roleDisplayName)($($roleId)) for resource $($roleResourceId) for the duration of $($minutesToAdd)($($duration)) minutes with the start time of $($dt.ToLocalTime().ToString("o"))($($startTime)) and the end time of end time $($endTimeString) for $($UserPrincipleName)($($UserAdObjectId)) and the reason is ""$($requestReason)""")
	}
	else
	{
		$processRoleRequest = $true
		$currentAttemp = 0
		
		while ($processRoleRequest)
		{
			if ($currentAttemp -lt $maxAttempts)
			{
				$Core.WriteCmEntry("Attemp ($($currentAttemp)/$($maxAttempts)) to request $($roleProvider) role $($roleDisplayName)($($roleId)) for resource $($roleResourceId) for the duration of $($minutesToAdd)($($duration)) minutes with the start time of $($dt.ToLocalTime().ToString("o")) and the end time of end time $($endTimeString) for $($UserPrincipleName)($($UserAdObjectId)) and the reason is ""$($requestReason)""")
				if ($roleProvider -eq 'AzureResources')
				{
					$guid = (New-Guid).Guid
					
					$requestResult = New-AzRoleAssignmentScheduleRequest -Name $guid -Scope $roleResourceId -ExpirationDuration "$($duration)" -ExpirationType AfterDuration -PrincipalId $UserAdObjectId -RequestType 'SelfActivate' -RoleDefinitionId $roleId -ScheduleInfoStartDateTime $startTime -Justification "$($requestreason)"
				}
				else
				{
					$requestResult = Open-AzureADMSPrivilegedRoleAssignmentRequest -ProviderId "$($roleProvider)" -ResourceId "$($roleResourceId)" -RoleDefinitionId "$($roleId)" -SubjectId "$($UserAdObjectId)" -Type 'UserAdd' -AssignmentState 'Active' -Schedule $schedule -Reason "$($requestreason)"
				}
				
				if (!$?)
				{
					$currentAttemp += 1
					if ($Error[0].Exception.Message.Contains('The time period of current request overlaps existing role assignment requests.'))
					{
						$Core.WriteCmEntry("Failed to request $($roleProvider) $($roleDisplayName) due to $($Error[0].Exception.Message), trying again in $($delayBetweenAttempts) seconds.")
						$Core.WriteCmEntry("Adjusting start and endtime for request $($roleProvider) $($roleDisplayName).")
						$dt = $dt.AddSeconds(60)
						$startTime = $dt.ToUniversalTime()
						$endTime = $dt.AddMinutes($minutesToAdd)
						$endTimeString = $endTime.ToLocalTime().ToString("o")
						$timeSpan = [TimeSpan]::FromTicks($endTime.Ticks - $dt.Ticks)
						$duration = [System.Xml.XmlConvert]::ToString($timeSpan)
					}
					elseif ($Error[0].Exception.Message.Contains('The following policy rules failed: ["MfaRule"]'))
					{
						$Core.WriteCmEntry("Failed to request $($roleProvider) $($roleDisplayName) due to $($Error[0].Exception.Message), trying again in $($delayBetweenAttempts) seconds.")
						$Core.WriteCmEntry("Fatal error $($Error[0].Exception.Message) exiting script")
						exit 5
					}
					elseif ($Error[0].Exception.Message.Contains('Role assignment already exists'))
					{
						$currentAttemp = $maxAttempts + 1
						$delayBetweenAttempts = 1
						$Core.WriteCmEntry("Failed to request $($roleProvider) $($roleDisplayName) due to $($Error[0].Exception.Message).")
						$Core.WriteCmEntry("Skipping for request $($roleProvider) $($roleDisplayName).")
					}
					Start-Sleep -Seconds $delayBetweenAttempts
				}
				else
				{
					$processRoleRequest = $false
					if ($DurationInMinutes -gt $MaxTimeAllowed)
					{
						RequestAzureRole -roleId $roleId -roleDisplayName $roleDisplayName -roleResourceId $roleResourceId -roleProvider $roleProvider -addStartTimeMinutes ($addStartTimeMinutes + $MaxTimeAllowed) -UserAdObjectId $UserAdObjectId -useWorkDay:$useWorkDay -requestReason $requestReason -UserPrincipleName $UserPrincipleName -MaxTimeAllowed $MaxTimeAllowed -DurationInMinutes ($DurationInMinutes - $MaxTimeAllowed) -endTime $endTime -startDay $startDay
					}
				}
			}
			else
			{
				$Core.WriteCmEntry("Unable to request $($roleProvider) $($roleDisplayName)($roleId) for $($UserPrincipleName)($($UserAdObjectId))")
				Write-Error "Unable to request $($roleProvider) $($roleDisplayName)($roleId) for $($UserPrincipleName)($($UserAdObjectId))"
				$processRoleRequest = $false
			}
		}
		
	}
}
$MyScriptVersion = "4.0.0.22"
$Core.WriteCmEntry("Script version: $($MyScriptVersion)")

$Core.WriteCmEntry("ScriptCore Version: $($Core.CoreVersion)")

$Core.WriteCmEntry("Script running under account: $($env:USERNAME)")

if ([string]::IsNullOrEmpty($UserPrincipleName))
{
	$Core.WriteCmEntry("The variable UserPrincipleName is empty, exiting.")
	exit 5
}

if ($Reason -eq '')
{
	$Core.WriteCmEntry("Reason is blank, exiting.")
	Stop-Transcript
	exit 6
}



if ($UseLatest)
{
	
	$tempItem = [RequiredModuleItem]::new("Az.Accounts")
	$Core.RequiredModuleItems.Add($tempItem)
	$tempItem = [RequiredModuleItem]::new("Az.Resources")
	$Core.RequiredModuleItems.Add($tempItem)
	$tempItem = [RequiredModuleItem]::new("AzureADPreview")
	$Core.RequiredModuleItems.Add($tempItem)
}
else
{
	$tempItem = [RequiredModuleItem]::new("Az.Accounts", "2.15.0")
	$Core.RequiredModuleItems.Add($tempItem)
	$tempItem = [RequiredModuleItem]::new("Az.Resources", "6.14.0")
	$Core.RequiredModuleItems.Add($tempItem)
	$tempItem = [RequiredModuleItem]::new("AzureADPreview", "2.0.2.183")
	$Core.RequiredModuleItems.Add($tempItem)
	
}
$Core.LoadRequiredModulesV2()

try
{
	$disconnectResultAzureAd = Disconnect-AzureAD -ErrorAction Stop
}
catch
{
	
}

try
{
	$disconnectResultAzAd = Disconnect-AzAccount -UserId $UserPrincipleName -ErrorAction Stop
}
catch
{
	
}

try
{
	Clear-AzContext -Force -ErrorAction Stop -Scope CurrentUser
}
catch
{
	
}

try
{
	Clear-AzContext -Force -ErrorAction Stop -Scope Process
}
catch
{
	
}

$Core.WriteCmEntry("UseWAM: $($UseWAM)")

$configValue = $UseWAM
$config = Get-AzConfig -EnableLoginByWam

if ($config.Value -ne $configValue)
{
	$config = Update-AzConfig -EnableLoginByWam $configValue.ToBool()
}
$Core.WriteCmEntry("Set-AZConfig -EnableLoginByWam is set to: $($config.Value)")

If (!$AzureAdConnection)
{
	try
	{
		Write-Log "Starting connection to AzureAd via the AzureADPreview module."
		$AzureAdConnection = Connect-AzureAD -ErrorAction Stop
		Write-Log "Connection established to AzureAd via the AzureADPreview module. Connected as $($AzureAdConnection.Account.Id) to $($AzureAdConnection.TenantDomain)($($AzureAdConnection.TenantId))"
	}
	catch
	{
		Write-LogError "Error connecting to AzureAd via the AzureADPreview module. Script exiting."
		Write-LogError "$($Error[0].Exception.Message)"
		exit
	}
	$myAd = Get-AzureADUser -ObjectId $AzureAdConnection.Account.Id
	$UserPrincipleName = $AzureAdConnection.Account.Id
	$Core.WriteCmEntry("Username to request roles for: $($UserPrincipleName)")
	
	$Core.WriteCmEntry("Setting clipboard text to: $($UserPrincipleName)")
	Set-Clipboard -Value "$($UserPrincipleName)" -Confirm:$false
	
	if (!$?)
	{
		Write-LogError "Error connecting to AzureAd via the AzureADPreview module. Script exiting."
		Write-LogError "$($Error[0].Exception.Message)"
		exit
	}
}
If (!$AzureAzConnection)
{
	try
	{
		Write-Log "Starting connection to AzureAd via the Az.Accounts module."
		
		$AzureAzConnection = Connect-AzAccount -ErrorAction Stop
		
		Write-Log "Connection established to AzureAd via the Az.Accounts module. Connected as $($AzureAzConnection.Context.Account.Id) to $($AzureAzConnection.Context.Tenant.Name)($($AzureAzConnection.Context.Tenant.Id))"
		
		$context = Get-AzContext -ErrorAction Stop
		
		$config = Get-AzConfig -DefaultSubscriptionForLogin
		
		if ($config.Value -ne $context.Subscription.Name)
		{
			$config = Update-AzConfig -DefaultSubscriptionForLogin $context.Subscription.Name
			$Core.WriteCmEntry("Set-AZConfig -DefaultSubscriptionForLogin to: $($config.Value)")
		}
		
	}
	catch
	{
		Write-LogError "Error connecting to AzureAd via the AzAccount module. Script exiting."
		Write-LogError "$($Error[0].Exception.Message)"
		exit
	}
	$tenantId = (Get-AzTenant | Where-Object { $_.Name -eq "Pactiv Evergreen" }).TenantId
	if (!$?)
	{
		Write-LogError "Error connecting to AzureAd via the AzAccount module. Script exiting."
		Write-LogError "$($Error[0].Exception.Message)"
		exit
	}
}

Write-Log "RequestRoles: $($RequestRoles)"
Write-Log "RequestResource: $($RequestResource)"

$RequestStartDay = [DateTime]::Parse($ScriptStartDay)
$RequestEndDay = [DateTime]::Parse($ScriptEndDay)

$WorkDaysToRequest = @()

$currentDay = [DateTime]::Now

if (($currentDay.DayOfWeek -eq 'Saturday' -or $currentDay.DayOfWeek -eq 'Sunday') -and $WorkWeek)
{
	$RequestStartDay = [System.DateTime]::Parse("12:45:00Z")
	$RequestEndDay = [System.DateTime]::Parse("12:55:00Z")
	$WorkWeek = $false
}

if (($RoleDurationInMinutes -gt 0) -and ($WorkDay -eq $false))
{
	$WorkDay = $false
	$WorkWeek = $false
	$RequestStartDay = [System.DateTime]::Parse("12:45:00Z")
	$RequestEndDay = [System.DateTime]::Parse("12:55:00Z")
}

Write-Log "WorkDay: $($WorkDay)"
Write-Log "WorkWeek: $($WorkWeek)"
Write-Log "RequestStartDay: $($RequestStartDay.ToString('o'))"
Write-Log "RequestEndDay: $($RequestEndDay.ToString('o'))"

if ($WorkWeek)
{
	do
	{
		$currentDay = [System.DateTime]::Parse("$($currentDay.ToShortDateString()) 12:45:00Z")
		if ($currentDay.DayOfYear -eq [DateTime]::Now.DayOfYear) { $currentDay = [DateTime]::Now }
		$WorkDaysToRequest += $currentDay
		$currentDay = $currentDay.AddDays(1)
	}
	until ($currentDay.DayOfWeek -eq 'Saturday' -or $currentDay.DayOfWeek -eq 'Sunday')
}
else
{
	$currentDay = $RequestStartDay
	do
	{
		if ($currentDay.DayOfWeek -eq 'Saturday' -or $currentDay.DayOfWeek -eq 'Sunday')
		{
			if (($currentDay.DayOfYear -eq $RequestEndDay.DayOfYear) -or $AllowWeekends)
			{
				$currentDay = [System.DateTime]::Parse("$($currentDay.ToShortDateString()) 12:45:00Z")
				if ($currentDay.DayOfYear -eq [DateTime]::Now.DayOfYear) { $currentDay = [DateTime]::Now }
				$WorkDaysToRequest += $currentDay
			}
		}
		else
		{
			$currentDay = [System.DateTime]::Parse("$($currentDay.ToShortDateString()) 12:45:00Z")
			if ($currentDay.DayOfYear -eq [DateTime]::Now.DayOfYear) { $currentDay = [DateTime]::Now }
			$WorkDaysToRequest += $currentDay
		}
		
		$currentDay = $currentDay.AddDays(1)
	}
	until ($currentDay.Ticks -ge $RequestEndDay.Ticks)
}

Write-Log "Requesting for the following dates: $($WorkDaysToRequest -join ', ')"
Write-Log ""

if ($RequestRoles)
{
	$roles = Get-AzureADMSPrivilegedRoleDefinition -ProviderId aadRoles -ResourceId $tenantId
	$myRoles = Get-AzureADMSPrivilegedRoleAssignment -ProviderId "aadRoles" -ResourceId $tenantId -Filter "subjectId eq '$($myad.ObjectId)'"
	$myEligibleRoles = $myRoles | Where-Object { $_.AssignmentState -eq 'Eligible' }
	
	if ($null -eq $RolesToRequest)
	{
		$RolesToRequest = @()
		foreach ($role in $myEligibleRoles)
		{
			$roleName = ($roles | Where-Object { $_.id -eq $role.RoleDefinitionId }).DisplayName
			$isActive = $myRoles | Where-Object { $_.RoleDefinitionId -eq $role.RoleDefinitionId -and $_.AssignmentState -eq 'Active' }
			$isExcluded = $RolesToExclude | Where-Object { $_ -eq $roleName }
			$Core.WriteCmEntryVerbose("Role: $($roleName), IsActive: $($Core.GetArrayCount($isActive)), IsExcluded: $($Core.GetArrayCount($isExcluded))")
			if ((Get-Date).DayOfYear -ne $RequestStartDay.DayOfYear) { $isActive = $null }
			if (($Core.GetArrayCount($isActive)) -eq 0 -and ($Core.GetArrayCount($isExcluded)) -eq 0)
			{
				$Core.WriteCmEntryVerbose("Adding role to activation list: $($roleName)")
				$RolesToRequest += $roleName
			}
			else
			{
				$Core.WriteCmEntry("Skipped: Adding role to activation list: $($roleName) as it's already active")
			}
		}
	}
	else
	{
		$Core.WriteCmEntry("The following Azure AAD Roles have been requested via the command line: $($RolesToRequest -join ', ')")
		$tempRoles = @()
		foreach ($role in $myEligibleRoles)
		{
			$roleName = ($roles | Where-Object { $_.id -eq $role.RoleDefinitionId }).DisplayName
			$isActive = $myRoles | Where-Object { $_.RoleDefinitionId -eq $role.RoleDefinitionId -and $_.AssignmentState -eq 'Active' }
			$isExcluded = $RolesToExclude | Where-Object { $_ -eq $roleName }
			$Core.WriteCmEntryVerbose("Role: $($roleName), IsActive: $($Core.GetArrayCount($isActive)), IsExcluded: $($Core.GetArrayCount($isExcluded))")
			if ((Get-Date).DayOfYear -ne $RequestStartDay.DayOfYear) { $isActive = $null }
			if (($Core.GetArrayCount($isActive)) -eq 0 -and ($Core.GetArrayCount($isExcluded)) -eq 0 -and $RolesToRequest -match $roleName)
			{
				$Core.WriteCmEntryVerbose("Adding role to activation list: $($roleName)")
				$tempRoles += $roleName
			}
			else
			{
				$Core.WriteCmEntry("Skipped: Adding role to activation list: $($roleName) as it's already active")
			}
		}
		$RolesToRequest = $tempRoles
	}
	
	$Core.WriteCmEntry("Roles to exclude: $($RolesToExclude -join ', ')")
	$Core.WriteCmEntry("Roles to activate: $($RolesToRequest -join ', ')")
	
	foreach ($currentDay in $WorkDaysToRequest)
	{
		Write-Log "Requesting roles for date: $($currentDay.ToLongDateString())"
		foreach ($roleName in $RolesToRequest)
		{
			if ($currentDay.DayOfYear -eq [DateTime]::Now.DayOfYear) { $currentDay = [DateTime]::Now }
			$role = $null
			$role = $roles | Where-Object { $_.DisplayName -eq $roleName }
			if ($role -eq $null)
			{
				$Core.WriteCmEntry("Skipping Azure AAD Role $($roleName) as the role doesn't exist.")
				continue
			}
			$roleSettings = Get-AzureADMSPrivilegedRoleSetting -ProviderId 'aadRoles' -Filter "ResourceId eq '$($tenantId)' and RoleDefinitionId eq '$($role.Id)'"
			$expirationRule = $roleSettings.UserMemberSettings | Where-Object { $_.RuleIdentifier -eq 'ExpirationRule' }
			$expirationSettings = $expirationRule.Setting | ConvertFrom-Json
			RequestAzureRole -roleId $role.Id -roleDisplayName $role.DisplayName -roleResourceId $role.ResourceId -roleProvider "aadRoles" -DurationInMinutes $RoleDurationInMinutes -MaxTimeAllowed $expirationSettings.maximumGrantPeriodInMinutes -UserAdObjectId $myAd.ObjectId -UserPrincipleName $myad.UserPrincipalName -useWorkDay:$WorkDay -requestReason $Reason -startDay $currentDay
		}
	}
}

if ($RequestResource)
{
	$eligibleAssignments = Get-AzRoleEligibilitySchedule -Scope "/" -Filter "asTarget()"
	$resourceRequests = @{ }
	$resourcesToActivate = @()
	if ($null -eq $ResourcesToRequest)
	{
		foreach ($assignment in $eligibleAssignments)
		{
			if ($resourceRequests.ContainsKey($assignment.Name) -eq $false)
			{
				$resourceRequests.Add($assignment.Name, $assignment)
				$resourcesToActivate += "$($assignment.RoleDefinitionDisplayName)"
			}
		}
	}
	else
	{
		foreach ($role in $ResourcesToRequest)
		{
			$assignments = $null
			$assignments = $eligibleAssignments | Where-Object { $_.RoleDefinitionDisplayName.ToLower().Trim() -eq $role.ToLower().Trim() }
			foreach ($assignment in $assignments)
			{
				if ($assignment -eq $null)
				{
					$Core.WriteCmEntry("Skipping requested Azure Resource Role $($role) as this role is not eligible for activation.")
					continue
				}
				
				if ($resourceRequests.ContainsKey($assignment.Name) -eq $false)
				{
					$resourceRequests.Add($assignment.Name, $assignment)
					$resourcesToActivate += "$($assignment.RoleDefinitionDisplayName)"
				}
			}
		}
	}
	$Core.WriteCmEntry("The following Azure Resource Roles have been requested: $($ResourcesToRequest -join ', ')")
	$Core.WriteCmEntry("Activating the following Azure Resources: $($resourcesToActivate -join ', ')")
	foreach ($currentDay in $WorkDaysToRequest)
	{
		Write-Log "Requesting resources for date: $($currentDay.ToLongDateString())"
		foreach ($assignment in $resourceRequests.Values)
		{
			if ($currentDay.DayOfYear -eq [DateTime]::Now.DayOfYear) { $currentDay = [DateTime]::Now }
			$Core.WriteCmEntry("Collecting information for $($assignment.RoleDefinitionDisplayName)($($assignment.Name)) for $($assignment.ResourceGroupName) in scope $($assignment.Scope)")
			$toSplit = $assignment.RoleDefinitionId -split '/'
			$assignment | Add-Member -MemberType NoteProperty -Name "RoleDefId" -Value $toSplit[-1] -Force
			$assignment | Add-Member -MemberType NoteProperty -Name "SubscriptionId" -Value $toSplit[2] -Force
			
			$managementPolicy = Get-AzRoleManagementPolicyAssignment -Scope $assignment.Scope | Where-Object { $_.roleDefinitionId -eq "$($assignment.RoleDefinitionId)" }
			if (!$?)
			{
				Start-Sleep -Seconds 5
				$managementPolicy = Get-AzRoleManagementPolicyAssignment -Scope $assignment.Scope | Where-Object { $_.roleDefinitionId -eq "$($assignment.RoleDefinitionId)" }
			}
			
			$assignment | Add-Member -MemberType NoteProperty -Name "ManagementPolicyName" -Value $managementPolicy.Name.Split('_')[0] -Force
			
			$managementPolicyDef = Get-AzRoleManagementPolicy -scope $assignment.Scope -Name $assignment.ManagementPolicyName
			if (!$?)
			{
				Start-Sleep -Seconds 5
				$managementPolicyDef = Get-AzRoleManagementPolicy -scope $assignment.Scope -Name $assignment.ManagementPolicyName
			}
			
			$ts = [System.Xml.XmlConvert]::ToTimeSpan(($managementPolicyDef.Rule | Where-Object { $_.Id -eq 'Expiration_EndUser_Assignment' }).MaximumDuration)
			$assignment | Add-Member -MemberType NoteProperty -Name "ManagementPolicy" -Value $managementPolicyDef -Force
			$assignment | Add-Member -MemberType NoteProperty -Name "MaximumMinutes" -Value $ts.TotalMinutes -Force
			
			RequestAzureRole -roleId $assignment.RoleDefinitionId -roleDisplayName $assignment.RoleDefinitionDisplayName -roleResourceId $assignment.Scope -roleProvider "azureResources" -DurationInMinutes $RoleDurationInMinutes -MaxTimeAllowed $assignment.MaximumMinutes -UserAdObjectId $myAd.ObjectId -UserPrincipleName $myad.UserPrincipalName -useWorkDay:$WorkDay -requestReason $Reason -startDay $currentDay
		}
	}
}

$Core.FinishScript()