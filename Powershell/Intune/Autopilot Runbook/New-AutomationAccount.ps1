
# Requires: Az.Accounts, Az.Automation, Microsoft.Graph modules
# Install-Module Az.Accounts,Az.Automation -Scope CurrentUser -Force
# Install-Module Microsoft.Graph -Scope CurrentUser -Force

param(
  [Parameter(Mandatory=$true)]
  [string]$TenantId,

  [Parameter(Mandatory=$true)]
  [string]$ResourceGroupName,

  [Parameter(Mandatory=$true)]
  [string]$AutomationAccountName
)

$PermissionName = "DeviceManagementServiceConfig.ReadWrite.All"

# 1) Get the managed identity objectId from the Automation Account (system-assigned)
#    This is the Entra ID service principal objectId you need for Graph assignment.
Connect-AzAccount -Tenant $TenantId | Out-Null

$aa = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction Stop
if (-not $aa.Identity -or -not $aa.Identity.PrincipalId) {
  throw "System-assigned managed identity is not enabled on Automation Account '$AutomationAccountName'. Enable it first."
}
$miObjectId = [Guid]$aa.Identity.PrincipalId
Write-Host "Managed Identity objectId: $miObjectId" -ForegroundColor Cyan

# 2) Connect to Microsoft Graph and assign the app role
Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All"
Select-MgProfile -Name v1.0

# Microsoft Graph service principal
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ConsistencyLevel eventual
if (-not $graphSp) { throw "Microsoft Graph service principal not found." }

# Find required app role (application permission)
$appRole = $graphSp.AppRoles | Where-Object {
  $_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains "Application"
}
if (-not $appRole) {
  throw "App role '$PermissionName' not found in Microsoft Graph. Verify the permission name."
}

# Idempotency check
$existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId -All |
  Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $graphSp.Id }

if ($existing) {
  Write-Host "SKIP: '$PermissionName' already assigned to $AutomationAccountName." -ForegroundColor Yellow
} else {
  $assignment = New-MgServicePrincipalAppRoleAssignment `
    -ServicePrincipalId $miObjectId `
    -PrincipalId $miObjectId `
    -ResourceId $graphSp.Id `
    -AppRoleId $appRole.Id

  Write-Host "SUCCESS: Granted '$PermissionName' to $AutomationAccountName ($miObjectId)." -ForegroundColor Green
}

Disconnect-MgGraph
