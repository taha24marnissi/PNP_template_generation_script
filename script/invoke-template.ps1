# Requires: PowerShell 7+, PnP.PowerShell module

param(
    [string]$TemplateFilePath = "C:\Users\tmarnissi\Desktop\sp pnp gen\test_microsoft_field.xml"
)

$SiteUrl = "https://3rd5pw.sharepoint.com/sites/mspatterntest001"
$SiteTitle = "mspatterntest001"
$SiteAlias = ""            # Optional. If empty, will be derived from $SiteUrl
$OwnerUpn = "taha@3rd5pw.onmicrosoft.com"

$ClientId = "a3164d9e-1c83-4ddc-a5c1-614135b82599"
$CertificatePath = "C:\Users\tmarnissi\Desktop\Sharepoint AI agent\sharepoint_Bot_Api\tahaUp.pfx"
$CertificatePassword = "1234"
$TenantDomain = "3rd5pw.onmicrosoft.com"
$TemplatePath = $TemplateFilePath
$ResourceFolder = ".\"
$AdminCenterUrl = "https://3rd5pw-admin.sharepoint.com"

$ErrorActionPreference = "Stop"

function Ensure-PnPModuleInstalled {
	if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
		Write-Host "Installing PnP.PowerShell for current user..." -ForegroundColor Yellow
		Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
	}
}

try {
	# Normalize to the script directory so relative paths resolve correctly
	$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
	Set-Location $scriptDir

	Ensure-PnPModuleInstalled

	if (-not (Test-Path -Path $TemplatePath)) {
		$resolved = Resolve-Path -LiteralPath $TemplatePath -ErrorAction SilentlyContinue
		throw "Template file not found at: $($resolved ?? $TemplatePath)"
	}

	if (-not (Test-Path -Path $ResourceFolder)) {
		Write-Host "Resource folder not found at $ResourceFolder. Proceeding without a resource folder." -ForegroundColor Yellow
		$ResourceFolder = $null
	}

	# Derive alias from SiteUrl when not provided; ensure SiteUrl aligns with alias
	if ([string]::IsNullOrWhiteSpace($SiteAlias)) {
		if ($SiteUrl -match "/sites/([^/]+)$") {
			$SiteAlias = $Matches[1]
		}
	}
	if (-not [string]::IsNullOrWhiteSpace($SiteAlias)) {
		$SiteUrl = "https://3rd5pw.sharepoint.com/sites/$SiteAlias"
	}

	$password = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force

	# Connect to admin center for site creation/check
	Write-Host "Connecting to admin center $AdminCenterUrl using certificate auth..." -ForegroundColor Cyan
	Connect-PnPOnline -Url $AdminCenterUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $password -Tenant $TenantDomain
	
	$scriptDir = "C:\Users\tmarnissi\Desktop\sp pnp gen\Scriptlogs"
    Set-Location $scriptDir

	$traceLogPath = Join-Path $scriptDir "TraceOutput-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"

	Start-PnPTraceLog -Path $traceLogPath -Level Debug

	# Ensure site exists
	$tenantSite = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
	if (-not $tenantSite) {
		if ([string]::IsNullOrWhiteSpace($SiteAlias)) {
			throw "SiteAlias is required to create a Microsoft 365 Group-connected team site. Set $SiteAlias or ensure $SiteUrl ends with /sites/<alias>."
		}
		Write-Host "Creating GROUP-connected site ($SiteTitle) at $SiteUrl..." -ForegroundColor Cyan
		New-PnPSite -Type TeamSite -Title $SiteTitle -Alias $SiteAlias -Owners $OwnerUpn
	}
	else {
		Write-Host "Site already exists at $SiteUrl" -ForegroundColor Yellow
	}

	# Wait until site is ready
	Write-Host "Waiting for site to be ready..." -ForegroundColor Cyan
	for ($i = 0; $i -lt 30; $i++) {
		$tenantSite = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
		if ($tenantSite -and $tenantSite.Status -eq "Active") { break }
		Start-Sleep -Seconds 10
	}
	if (-not $tenantSite -or $tenantSite.Status -ne "Active") {
		throw "Site did not become active in time. Current status: $($tenantSite.Status)"
	}

	# Connect to the newly created/existing site and apply the site template
	Write-Host "Connecting to $SiteUrl using certificate auth..." -ForegroundColor Cyan
	Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $password -Tenant $TenantDomain

	Write-Host "Applying provisioning template from $TemplatePath" -ForegroundColor Cyan
	if ($null -ne $ResourceFolder) {
		Invoke-PnPSiteTemplate -Path $TemplatePath -ResourceFolder $ResourceFolder
	}
	else {
		Invoke-PnPSiteTemplate -Path $TemplatePath
	}

	Write-Host "Provisioning completed successfully." -ForegroundColor Green
}
catch {
	Write-Error $_
	exit 1
}
finally {
	try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}