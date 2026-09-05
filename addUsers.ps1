[CmdletBinding(SupportsShouldProcess)]
param(
	[ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
	[string]$CsvPath = (Join-Path $PSScriptRoot 'users.csv'),

	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[string]$TeamName,

	[ValidatePattern('^[A-Z]{2}$')]
	[string]$UsageLocation = 'NL',

	[char]$CsvDelimiter = ','
)

# todo, ask ai to create team, and create assigment with due date and assign automatically to current and new users.
# probably also assign the teachers/owners.

# run command for highger version powershell: pwsh .\addUsers.ps1 -TeamName "175 - TILBURG"

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$requiredScopes = @(
	'GroupMember.ReadWrite.All'
	'LicenseAssignment.ReadWrite.All'
	'Organization.Read.All'
	'User.ReadWrite.All'
)
Connect-MgGraph -Scopes $requiredScopes -NoWelcome

function Get-GraphCollection {
	param([Parameter(Mandatory)][string]$Uri)

	$items = [System.Collections.Generic.List[object]]::new()
	while ($Uri) {
		$response = Invoke-MgGraphRequest -Method GET -Uri $Uri
		foreach ($item in $response.value) {
			$items.Add($item)
		}
		$Uri = $response.'@odata.nextLink'
	}
	return $items
}

function Get-EntraUser {
	param([Parameter(Mandatory)][string]$UserPrincipalName)

	$encodedUserPrincipalName = [Uri]::EscapeDataString($UserPrincipalName)
	try {
		return Invoke-MgGraphRequest -Method GET -Uri "/v1.0/users/$encodedUserPrincipalName`?`$select=id,userPrincipalName,otherMails,assignedLicenses"
	}
	catch {
		if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
			return $null
		}
		throw
	}
}

$rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter $CsvDelimiter)
if ($rows.Count -eq 0) {
	throw "CSV file '$CsvPath' contains no users."
}

$requiredColumns = @('username', 'password', 'privateEmail', 'license')
$columnNames = @($rows[0].PSObject.Properties.Name)
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $columnNames })
if ($missingColumns.Count -gt 0) {
	throw "CSV is missing required column(s): $($missingColumns -join ', ')."
}

$groups = @(Get-GraphCollection -Uri '/v1.0/groups?$select=id,displayName,resourceProvisioningOptions')
$teams = @($groups | Where-Object { 'Team' -in $_.resourceProvisioningOptions })
$teamMatches = @($teams | Where-Object { $_.displayName -ieq $TeamName })
if ($teamMatches.Count -ne 1) {
	$visibleTeamNames = @($teams.displayName | Sort-Object | Select-Object -First 25)
	$visibleTeamsMessage = if ($visibleTeamNames.Count) {
		" Teams visible to Graph include: $($visibleTeamNames -join ', ')."
	}
	else {
		' Graph did not return any Teams.'
	}
	throw "Expected exactly one Team named '$TeamName', but found $($teamMatches.Count).$visibleTeamsMessage"
}
$team = $teamMatches[0]

$skus = @(Get-GraphCollection -Uri '/v1.0/subscribedSkus?$select=skuId,skuPartNumber,capabilityStatus,consumedUnits,prepaidUnits')

$memberIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$members = Get-GraphCollection -Uri "/v1.0/groups/$($team.id)/members?`$select=id"
foreach ($member in $members) {
	[void]$memberIds.Add([string]$member.id)
}
$ownerIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$owners = Get-GraphCollection -Uri "/v1.0/groups/$($team.id)/owners?`$select=id"
foreach ($owner in $owners) {
	[void]$ownerIds.Add([string]$owner.id)
}

$results = foreach ($row in $rows) {
	$username = ([string]$row.username).Trim()
	$password = [string]$row.password
	$privateEmail = ([string]$row.privateEmail).Trim()
	$skuPartNumber = ([string]$row.license).Trim()

	try {
		if (-not $username -or -not $skuPartNumber) {
			throw 'username and license must both have a value.'
		}
		if ($username -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
			throw 'username must be a valid email address.'
		}
		$isFaculty = $skuPartNumber -eq 'STANDARDWOFFPACK_FACULTY'
		if (-not $isFaculty -and $skuPartNumber -ne 'STANDARDWOFFPACK_STUDENT') {
			throw "Unsupported license '$skuPartNumber'. Use STANDARDWOFFPACK_STUDENT or STANDARDWOFFPACK_FACULTY."
		}

		$skuMatches = @($skus | Where-Object { $_.skuPartNumber -eq $skuPartNumber })
		if ($skuMatches.Count -ne 1) {
			throw "Expected exactly one subscribed SKU '$skuPartNumber', but found $($skuMatches.Count)."
		}
		$sku = $skuMatches[0]
		if ($sku.capabilityStatus -ne 'Enabled') {
			throw "Subscribed SKU '$skuPartNumber' is not enabled."
		}

		$user = Get-EntraUser -UserPrincipalName $username
		$userAlreadyExists = $null -ne $user
		if (-not $user) {
			if (-not $password -or -not $privateEmail) {
				throw 'password and privateEmail must both have a value when creating a new user.'
			}
			if ($privateEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
				throw 'privateEmail must be a valid email address.'
			}

			$mailNickname = ($username.Split('@')[0] -replace '[^a-zA-Z0-9._-]', '')
			$createBody = @{
				accountEnabled    = $true
				displayName       = $username
				mailNickname      = $mailNickname
				otherMails        = @($privateEmail)
				passwordProfile   = @{
					forceChangePasswordNextSignIn = $true
					password                      = $password
				}
				usageLocation     = $UsageLocation.ToUpperInvariant()
				userPrincipalName = $username
			}

			if ($PSCmdlet.ShouldProcess($username, 'Create Entra user')) {
				$user = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/users' -Body ($createBody | ConvertTo-Json -Depth 5)
			}
		}

		if (-not $userAlreadyExists -and $user -and $sku.skuId -notin @($user.assignedLicenses.skuId)) {
			$licenseBody = @{
				addLicenses    = @(@{ disabledPlans = @(); skuId = $sku.skuId })
				removeLicenses = @()
			}
			if ($PSCmdlet.ShouldProcess($username, "Assign license $skuPartNumber")) {
				Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$($user.id)/assignLicense" -Body ($licenseBody | ConvertTo-Json -Depth 5) | Out-Null
			}
		}

		if ($user -and -not $memberIds.Contains([string]$user.id)) {
			$memberBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" }
			if ($PSCmdlet.ShouldProcess($username, "Add as member to Team $TeamName")) {
				Invoke-MgGraphRequest -Method POST -Uri "/v1.0/groups/$($team.id)/members/`$ref" -Body ($memberBody | ConvertTo-Json) | Out-Null
				[void]$memberIds.Add([string]$user.id)
			}
		}

		if ($isFaculty -and $user -and -not $ownerIds.Contains([string]$user.id)) {
			$ownerBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" }
			if ($PSCmdlet.ShouldProcess($username, "Add as owner to Team $TeamName")) {
				Invoke-MgGraphRequest -Method POST -Uri "/v1.0/groups/$($team.id)/owners/`$ref" -Body ($ownerBody | ConvertTo-Json) | Out-Null
				[void]$ownerIds.Add([string]$user.id)
			}
		}

		[pscustomobject]@{
			Username = $username
			TeamRole = if ($isFaculty) { 'Owner' } else { 'Member' }
			Status   = if ($WhatIfPreference) { 'WhatIf' } else { 'Success' }
			Message  = if ($userAlreadyExists) { 'Existing Entra user; profile and license unchanged.' } else { '' }
		}
	}
	catch {
		[pscustomobject]@{
			Username = $username
			Status   = 'Failed'
			Message  = $_.Exception.Message
		}
	}
}

$results | Format-Table -AutoSize
if ($results.Status -contains 'Failed') {
	throw 'One or more users could not be processed. See the results above.'
}
