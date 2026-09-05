[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[string]$TeamName,

	[ValidateNotNullOrEmpty()]
	[string]$Description = 'Class Team created from the Microsoft Teams education class template.',

	[Parameter(Mandatory)]
	[datetime]$DueDateTime,

	[ValidateNotNullOrEmpty()]
	[string]$AssignmentName = 'OptodermGP Casuïstiek',

	[ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
	[string]$HomeworkPath = (Join-Path $PSScriptRoot 'Huiswerk.pptx'),

	[ValidateNotNullOrEmpty()]
	[string]$TeacherLibraryName = 'Docenten Documenten',

	[string]$PnPClientId = $(
		if ($env:ENTRAID_APP_ID) { $env:ENTRAID_APP_ID }
		elseif ($env:ENTRAID_CLIENT_ID) { $env:ENTRAID_CLIENT_ID }
		elseif ($env:AZURE_CLIENT_ID) { $env:AZURE_CLIENT_ID }
		else { 'd1887c5e-e6f6-415c-9b07-430abc83901b' }
	)
)

$ErrorActionPreference = 'Stop'

# run command: pwsh .\createCourse.ps1 -TeamName "173 - HOLTEN" -DueDateTime "2026-10-23 23:59"

if ($DueDateTime -le (Get-Date)) {
	throw 'DueDateTime must be in the future.'
}

if ([string]::IsNullOrWhiteSpace($PnPClientId)) {
	throw 'PnPClientId is required. Pass -PnPClientId or set ENTRAID_APP_ID to an interactive Entra app with delegated SharePoint AllSites.FullControl permission.'
}

if (-not $PSCmdlet.ShouldProcess($TeamName, "Create Microsoft Teams education class and publish assignment '$AssignmentName'")) {
	return
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module PnP.PowerShell -ErrorAction Stop
$graphScopes = @(
	'Channel.ReadBasic.All'
	'EduAssignments.ReadWrite'
	'Files.ReadWrite.All'
	'Sites.Read.All'
	'Team.Create'
	'TeamsTab.ReadWriteForTeam'
)
Connect-MgGraph -Scopes $graphScopes -NoWelcome
$graphContext = Get-MgContext
if ($graphContext.Scopes -notcontains 'Files.ReadWrite.All') {
	throw "The Microsoft Graph token for '$($graphContext.Account)' does not contain Files.ReadWrite.All. Disconnect from Graph and sign in again to grant access to shared Team files."
}

function Wait-TeamProvisioningOperation {
	param(
		[Parameter(Mandatory)]
		[string]$OperationLocation,

		[Parameter(Mandatory)]
		[string]$TeamName,

		[timespan]$Timeout = [timespan]::FromMinutes(10)
	)

	$deadline = (Get-Date).Add($Timeout)
	do {
		$operation = $null
		try {
			$operation = Invoke-MgGraphRequest -Method GET -Uri $OperationLocation
		}
		catch {
			$errorText = "$($_.Exception.Message)`n$($_.ErrorDetails.Message)"
			$isTransientOperationLookup =
				$errorText -match '404 Not Found' -and
				$errorText -match 'Operation id not found for given Team|No workflow found\s+with\s+supplied\s+ID'
			if (-not $isTransientOperationLookup -or (Get-Date) -ge $deadline) {
				throw
			}

			Write-Warning 'The Team provisioning operation is not visible yet. Retrying in 5 seconds.'
		}

		if ($operation.status -eq 'failed') {
			throw "Team provisioning failed: $($operation.error.message)"
		}
		if ($operation.status -eq 'succeeded') {
			return $operation
		}
		if ((Get-Date) -ge $deadline) {
			throw "Timed out waiting for Team '$TeamName' to finish provisioning."
		}

		Start-Sleep -Seconds 5
	} while ($true)
}

function Get-TeamSiteUrl {
	param(
		[Parameter(Mandatory)]
		[string]$TeamId,

		[timespan]$Timeout = [timespan]::FromMinutes(10)
	)

	$deadline = (Get-Date).Add($Timeout)
	do {
		try {
			$site = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups/$TeamId/sites/root"
			if ($site.webUrl) {
				return [string]$site.webUrl
			}
		}
		catch {
			if ((Get-Date) -ge $deadline) {
				throw
			}
		}

		if ((Get-Date) -ge $deadline) {
			throw "Timed out waiting for the SharePoint site for Team '$TeamId'."
		}
		Start-Sleep -Seconds 10
	} while ($true)
}

function Set-UpAssignmentResourcesFolder {
	param(
		[Parameter(Mandatory)]
		[string]$ClassId,

		[Parameter(Mandatory)]
		[string]$AssignmentId,

		[timespan]$Timeout = [timespan]::FromMinutes(10)
	)

	$deadline = (Get-Date).Add($Timeout)
	do {
		try {
			return Invoke-MgGraphRequest `
				-Method POST `
				-Uri "/v1.0/education/classes/$ClassId/assignments/$AssignmentId/setUpResourcesFolder" `
				-Body '{}'
		}
		catch {
			$errorText = "$($_.Exception.Message)`n$($_.ErrorDetails.Message)"
			if ($errorText -notmatch 'classNotSetUpInSharePoint' -or (Get-Date) -ge $deadline) {
				throw
			}

			Write-Warning 'The education class is still being connected to SharePoint. Retrying assignment resource setup in 10 seconds.'
			Start-Sleep -Seconds 10
		}
	} while ($true)
}

function Set-TeacherDocumentsTab {
	param(
		[Parameter(Mandatory)]
		[string]$TeamId,

		[Parameter(Mandatory)]
		[string]$LibraryUrl,

		[Parameter(Mandatory)]
		[string]$DisplayName,

		[timespan]$Timeout = [timespan]::FromMinutes(10)
	)

	$deadline = (Get-Date).Add($Timeout)
	do {
		try {
			$channel = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/teams/$TeamId/primaryChannel"
			break
		}
		catch {
			$errorText = "$($_.Exception.Message)`n$($_.ErrorDetails.Message)"
			if ($errorText -notmatch '404|NotFound|ItemNotFound|GetThreadS2SRequest' -or (Get-Date) -ge $deadline) {
				throw
			}
			Write-Warning "The Team primary channel is still propagating. Retrying the '$DisplayName' tab in 10 seconds."
			Start-Sleep -Seconds 10
		}
	} while ($true)

	$channelId = [Uri]::EscapeDataString([string]$channel.id)
	$tabsUri = "/v1.0/teams/$TeamId/channels/$channelId/tabs"
	$tabs = @((Invoke-MgGraphRequest -Method GET -Uri "$tabsUri`?`$expand=teamsApp").value)
	$teacherDocumentsTab = $tabs | Where-Object {
		$_.displayName -eq $DisplayName -or $_.configuration.contentUrl -eq $LibraryUrl
	} | Select-Object -First 1
	$websiteAppId = 'com.microsoft.teamspace.tab.web'

	if ($teacherDocumentsTab -and (
		$teacherDocumentsTab.teamsApp.id -ne $websiteAppId -or
		$teacherDocumentsTab.configuration.contentUrl -ne $LibraryUrl)) {
		Invoke-MgGraphRequest -Method DELETE -Uri "$tabsUri/$($teacherDocumentsTab.id)"
		$teacherDocumentsTab = $null
	}

	if (-not $teacherDocumentsTab) {
		$tabBody = @{
			displayName           = $DisplayName
			'teamsApp@odata.bind' = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$websiteAppId"
			configuration         = @{
				entityId   = $LibraryUrl
				contentUrl  = $LibraryUrl
				websiteUrl  = $LibraryUrl
				removeUrl   = $null
			}
		}
		$teacherDocumentsTab = Invoke-MgGraphRequest -Method POST -Uri $tabsUri -Body ($tabBody | ConvertTo-Json -Depth 5) -ContentType 'application/json'
	}
	elseif ($teacherDocumentsTab.displayName -ne $DisplayName) {
		$tabBody = @{ displayName = $DisplayName }
		$teacherDocumentsTab = Invoke-MgGraphRequest -Method PATCH -Uri "$tabsUri/$($teacherDocumentsTab.id)" -Body ($tabBody | ConvertTo-Json) -ContentType 'application/json'
	}

	$verifiedTab = Invoke-MgGraphRequest -Method GET -Uri "$tabsUri/$($teacherDocumentsTab.id)`?`$expand=teamsApp"
	if ($verifiedTab.displayName -ne $DisplayName -or
		$verifiedTab.teamsApp.id -ne $websiteAppId -or
		$verifiedTab.configuration.contentUrl -ne $LibraryUrl) {
		throw "'$DisplayName' tab verification failed for Team '$TeamId'."
	}

	$verifiedTab
}

function Send-AssignmentResourceFile {
	param(
		[Parameter(Mandatory)]
		[string]$UploadUri,

		[Parameter(Mandatory)]
		[string]$FilePath,

		[timespan]$Timeout = [timespan]::FromMinutes(10)
	)

	$deadline = (Get-Date).Add($Timeout)
	do {
		try {
			return Invoke-MgGraphRequest `
				-Method PUT `
				-Uri $UploadUri `
				-InputFilePath $FilePath `
				-ContentType 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
		}
		catch {
			$errorText = "$($_.Exception.Message)`n$($_.ErrorDetails.Message)"
			if ($errorText -notmatch 'accessDenied|Access denied' -or (Get-Date) -ge $deadline) {
				throw
			}

			Write-Warning 'The assignment folder permissions are still propagating in SharePoint. Retrying the file upload in 10 seconds.'
			Start-Sleep -Seconds 10
		}
	} while ($true)
}

function New-BlankExcelWorkbookStream {
	$stream = [System.IO.MemoryStream]::new()
	$archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
	$parts = [ordered]@{
		'[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
		'_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
		'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Status" sheetId="1" r:id="rId1"/></sheets></workbook>'
		'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
		'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
	}

	try {
		foreach ($part in $parts.GetEnumerator()) {
			$entry = $archive.CreateEntry($part.Key, [System.IO.Compression.CompressionLevel]::Optimal)
			$writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
			try {
				$writer.Write($part.Value)
			}
			finally {
				$writer.Dispose()
			}
		}
	}
	finally {
		$archive.Dispose()
	}

	$stream.Position = 0
	$stream
}

function Add-StatusWorkbook {
	param(
		[Parameter(Mandatory)]
		[object]$Connection,

		[Parameter(Mandatory)]
		[string]$SiteUrl,

		[Parameter(Mandatory)]
		[string]$FolderUrl
	)

	$fileName = 'Status.xlsx'
	$filePath = "$($FolderUrl.TrimEnd('/'))/$fileName"
	$file = Get-PnPFile -Url $filePath -AsFileObject -Connection $Connection -ErrorAction SilentlyContinue
	if (-not $file) {
		$workbookStream = New-BlankExcelWorkbookStream
		try {
			$file = Add-PnPFile -FileName $fileName -Folder $FolderUrl -Stream $workbookStream -Connection $Connection
		}
		finally {
			$workbookStream.Dispose()
		}
	}

	if (-not $file) {
		throw "Failed to create '$fileName' in '$FolderUrl'."
	}

	$siteUri = [Uri]$SiteUrl
	"$($siteUri.Scheme)://$($siteUri.Authority)$filePath"
}

function New-OwnersOnlyDocumentLibrary {
	param(
		[Parameter(Mandatory)]
		[string]$SiteUrl,

		[Parameter(Mandatory)]
		[string]$LibraryName,

		[Parameter(Mandatory)]
		[string]$ClientId
	)

	$connection = Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ReturnConnection
	$library = Get-PnPList -Identity $LibraryName -Connection $connection -ErrorAction SilentlyContinue
	if (-not $library) {
		$libraryUrl = $LibraryName -replace '[^A-Za-z0-9_-]', ''
		if ([string]::IsNullOrWhiteSpace($libraryUrl)) {
			throw "Library name '$LibraryName' does not contain any characters that can be used in its URL."
		}
		$library = New-PnPList -Title $LibraryName -Url $libraryUrl -Template DocumentLibrary -EnableVersioning -OnQuickLaunch -Connection $connection
	}

	$web = Get-PnPWeb -Includes AssociatedOwnerGroup -Connection $connection
	$ownersGroup = $web.AssociatedOwnerGroup
	if (-not $ownersGroup) {
		throw "SharePoint site '$SiteUrl' does not have an associated Owners group."
	}

	Get-PnPProperty -ClientObject $library -Property HasUniqueRoleAssignments -Connection $connection | Out-Null
	if (-not $library.HasUniqueRoleAssignments) {
		Set-PnPList -Identity $library -BreakRoleInheritance -CopyRoleAssignments -Connection $connection | Out-Null
	}

	$library = Get-PnPList -Identity $LibraryName -Connection $connection
	Get-PnPProperty -ClientObject $library -Property RoleAssignments -Connection $connection | Out-Null
	$roleAssignments = @($library.RoleAssignments)
	$ownersAssignment = $false
	foreach ($roleAssignment in $roleAssignments) {
		Get-PnPProperty -ClientObject $roleAssignment -Property Member -Connection $connection | Out-Null
		if ($roleAssignment.Member.Id -eq $ownersGroup.Id) {
			$ownersAssignment = $true
		}
	}
	if (-not $ownersAssignment) {
		throw "The associated Owners group '$($ownersGroup.Title)' does not have access to '$LibraryName'. No permissions were removed."
	}

	foreach ($roleAssignment in $roleAssignments) {
		if ($roleAssignment.Member.Id -ne $ownersGroup.Id) {
			$roleAssignment.DeleteObject()
		}
	}
	Invoke-PnPQuery -Connection $connection | Out-Null

	$library = Get-PnPList -Identity $LibraryName -Connection $connection
	Get-PnPProperty -ClientObject $library -Property HasUniqueRoleAssignments, RoleAssignments, RootFolder -Connection $connection | Out-Null
	$remainingPrincipals = foreach ($roleAssignment in @($library.RoleAssignments)) {
		Get-PnPProperty -ClientObject $roleAssignment -Property Member -Connection $connection | Out-Null
		$roleAssignment.Member
	}
	if (-not $library.HasUniqueRoleAssignments -or @($remainingPrincipals).Count -ne 1 -or $remainingPrincipals[0].Id -ne $ownersGroup.Id) {
		throw "Permission verification failed for '$LibraryName'. Expected only '$($ownersGroup.Title)' to retain access."
	}

	$libraryAbsoluteUrl = "{0}://{1}{2}" -f ([Uri]$SiteUrl).Scheme, ([Uri]$SiteUrl).Authority, $library.RootFolder.ServerRelativeUrl

	[pscustomobject]@{
		Name           = $library.Title
		Url            = $libraryAbsoluteUrl
		OwnersGroup    = $ownersGroup.Title
		UniqueSecurity = $library.HasUniqueRoleAssignments
		FolderUrl      = $library.RootFolder.ServerRelativeUrl
		Connection     = $connection
	}
}

$requestBody = @{
	'template@odata.bind' = "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')"
	displayName           = $TeamName
	description           = $Description
}

$responseHeaders = $null
$statusCode = $null

Invoke-MgGraphRequest `
	-Method POST `
	-Uri '/v1.0/teams' `
	-Body ($requestBody | ConvertTo-Json) `
	-ContentType 'application/json' `
	-ResponseHeadersVariable responseHeaders `
	-StatusCodeVariable statusCode | Out-Null

if ($statusCode -ne 202) {
	throw "Microsoft Graph returned unexpected status code $statusCode while creating '$TeamName'."
}

$operationLocation = [string]$responseHeaders.Location
if (-not $operationLocation) {
	throw 'Microsoft Graph did not return a Team provisioning operation location.'
}
if ($operationLocation -match '^https://graph\.microsoft\.com/(?!(?:v1\.0|beta)/)') {
	$operationLocation = $operationLocation -replace '^https://graph\.microsoft\.com/', 'https://graph.microsoft.com/v1.0/'
}
elseif ($operationLocation -match '^/(?!v1\.0/|beta/)') {
	$operationLocation = "/v1.0$operationLocation"
}

$operation = Wait-TeamProvisioningOperation `
	-OperationLocation $operationLocation `
	-TeamName $TeamName

$teamId = [string]$operation.targetResourceId
if (-not $teamId -and [string]$responseHeaders.'Content-Location' -match "teams\('(?<TeamId>[^']+)'\)") {
	$teamId = $Matches.TeamId
}
if (-not $teamId) {
	throw 'Microsoft Graph did not return the ID of the provisioned Team.'
}

$teamSiteUrl = Get-TeamSiteUrl -TeamId $teamId
$teacherLibrary = New-OwnersOnlyDocumentLibrary `
	-SiteUrl $teamSiteUrl `
	-LibraryName $TeacherLibraryName `
	-ClientId $PnPClientId
$statusWorkbookUrl = Add-StatusWorkbook `
	-Connection $teacherLibrary.Connection `
	-SiteUrl $teamSiteUrl `
	-FolderUrl $teacherLibrary.FolderUrl

$assignmentBody = @{
	displayName          = $AssignmentName
	dueDateTime          = $DueDateTime.ToUniversalTime().ToString('o')
	assignTo             = @{
		'@odata.type' = '#microsoft.graph.educationAssignmentClassRecipient'
	}
	addedStudentAction   = 'assignIfOpen'
	allowLateSubmissions = $true
	status               = 'draft'
}

$assignment = Invoke-MgGraphRequest `
	-Method POST `
	-Uri "/v1.0/education/classes/$teamId/assignments" `
	-Body ($assignmentBody | ConvertTo-Json -Depth 5) `
	-ContentType 'application/json'

$assignmentWithFolder = Set-UpAssignmentResourcesFolder `
	-ClassId $teamId `
	-AssignmentId $assignment.id

$resourcesFolderUrl = [string]$assignmentWithFolder.resourcesFolderUrl
if ([string]::IsNullOrWhiteSpace($resourcesFolderUrl)) {
	throw 'Microsoft Graph did not return an assignment resources folder.'
}

$homeworkFile = Get-Item -LiteralPath $HomeworkPath
$encodedFileName = [Uri]::EscapeDataString($homeworkFile.Name)
$uploadedFile = Send-AssignmentResourceFile `
	-UploadUri "${resourcesFolderUrl}:/${encodedFileName}:/content" `
	-FilePath $homeworkFile.FullName

$fileUrl = "https://graph.microsoft.com/v1.0/drives/$($uploadedFile.parentReference.driveId)/items/$($uploadedFile.id)"
$resourceBody = @{
	distributeForStudentWork = $true
	resource                 = @{
		'@odata.type' = '#microsoft.graph.educationPowerPointResource'
		displayName   = $homeworkFile.Name
		fileUrl       = $fileUrl
	}
}

$assignmentResource = Invoke-MgGraphRequest `
	-Method POST `
	-Uri "/v1.0/education/classes/$teamId/assignments/$($assignment.id)/resources" `
	-Body ($resourceBody | ConvertTo-Json -Depth 5) `
	-ContentType 'application/json'

$publishedAssignment = Invoke-MgGraphRequest `
	-Method POST `
	-Uri "/v1.0/education/classes/$teamId/assignments/$($assignment.id)/publish"

$teacherDocumentsTab = Set-TeacherDocumentsTab `
	-TeamId $teamId `
	-LibraryUrl $teacherLibrary.Url `
	-DisplayName $TeacherLibraryName

[pscustomobject]@{
	TeamName          = $TeamName
	TeamId            = $teamId
	TeamStatus        = 'Provisioned'
	AssignmentName    = $publishedAssignment.displayName
	AssignmentId      = $publishedAssignment.id
	AssignmentStatus  = $publishedAssignment.status
	AssignmentUrl     = $publishedAssignment.webUrl
	HomeworkFile      = $assignmentResource.resource.displayName
	StudentCopy       = $assignmentResource.distributeForStudentWork
	SharePointSiteUrl = $teamSiteUrl
	TeacherLibrary    = $teacherLibrary.Name
	TeacherLibraryUrl = $teacherLibrary.Url
	LibraryOwners     = $teacherLibrary.OwnersGroup
	UniquePermissions = $teacherLibrary.UniqueSecurity
	StatusWorkbookUrl = $statusWorkbookUrl
	TeacherTabUrl     = $teacherDocumentsTab.configuration.contentUrl
	OperationLocation = $operationLocation
	TeamLocation      = [string]$responseHeaders.'Content-Location'
}
