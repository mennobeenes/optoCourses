[CmdletBinding(SupportsShouldProcess)]
param(
	[ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
	[string]$CsvPath = (Join-Path $PSScriptRoot 'users.csv'),

	[string]$TeamName = '173 - HOLTEN',

	[string]$SkuPartNumber = 'STANDARDWOFFPACK_STUDENT',

	[ValidatePattern('^[A-Z]{2}$')]
	[string]$UsageLocation = 'NL',

	[char]$CsvDelimiter = ','
)

# todo, ask ai to create team, and create assigment with due date and assign automatically to current and new users.
# probably also assign the teachers/owners.

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$requiredScopes = @(
	'GroupMember.ReadWrite.All'
	'LicenseAssignment.ReadWrite.All'
	'Organization.Read.All'
	'User.ReadWrite.All'
)
Connect-MgGraph -Scopes $requiredScopes -UseDeviceCode -NoWelcome

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

$requiredColumns = @('username', 'password', 'privateEmail')
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
$skuMatches = @($skus | Where-Object { $_.skuPartNumber -eq $SkuPartNumber })
if ($skuMatches.Count -ne 1) {
	throw "Expected exactly one subscribed SKU '$SkuPartNumber', but found $($skuMatches.Count)."
}
$sku = $skuMatches[0]
if ($sku.capabilityStatus -ne 'Enabled') {
	throw "Subscribed SKU '$SkuPartNumber' is not enabled."
}

$memberIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$members = Get-GraphCollection -Uri "/v1.0/groups/$($team.id)/members?`$select=id"
foreach ($member in $members) {
	[void]$memberIds.Add([string]$member.id)
}

$results = foreach ($row in $rows) {
	$username = ([string]$row.username).Trim()
	$password = [string]$row.password
	$privateEmail = ([string]$row.privateEmail).Trim()

	try {
		if (-not $username -or -not $password -or -not $privateEmail) {
			throw 'username, password, and privateEmail must all have a value.'
		}
		if ($username -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or $privateEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
			throw 'username and privateEmail must be valid email addresses.'
		}
		$user = Get-EntraUser -UserPrincipalName $username
		if (-not $user) {
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
		else {
			$remainingEmails = @($user.otherMails | Where-Object { $_ -and $_ -ine $privateEmail })
			$updateBody = @{
				otherMails    = @($privateEmail) + $remainingEmails
				usageLocation = $UsageLocation.ToUpperInvariant()
			}
			if ($PSCmdlet.ShouldProcess($username, 'Update private email and usage location')) {
				Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/users/$($user.id)" -Body ($updateBody | ConvertTo-Json) | Out-Null
			}
		}

		if ($user -and $sku.skuId -notin @($user.assignedLicenses.skuId)) {
			$licenseBody = @{
				addLicenses    = @(@{ disabledPlans = @(); skuId = $sku.skuId })
				removeLicenses = @()
			}
			if ($PSCmdlet.ShouldProcess($username, "Assign license $SkuPartNumber")) {
				Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$($user.id)/assignLicense" -Body ($licenseBody | ConvertTo-Json -Depth 5) | Out-Null
			}
		}

		if ($user -and -not $memberIds.Contains([string]$user.id)) {
			$memberBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" }
			if ($PSCmdlet.ShouldProcess($username, "Add to Team $TeamName")) {
				Invoke-MgGraphRequest -Method POST -Uri "/v1.0/groups/$($team.id)/members/`$ref" -Body ($memberBody | ConvertTo-Json) | Out-Null
				[void]$memberIds.Add([string]$user.id)
			}
		}

		[pscustomobject]@{
			Username = $username
			Status   = if ($WhatIfPreference) { 'WhatIf' } else { 'Success' }
			Message  = ''
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
