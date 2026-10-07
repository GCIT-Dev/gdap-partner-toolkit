#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Creates an app registration that can read or write one SharePoint site through Microsoft Graph, using
    Sites.Selected and a certificate.

.DESCRIPTION
    Some scripts need unattended access to a SharePoint site, for example to keep a list up to date. The
    least privileged way to give an app that access is the Sites.Selected application permission, which on
    its own grants nothing, plus a permission on the one site the app needs. This script, run against one
    tenant (usually your own, or a customer's through GDAP):

    1. resolves -SiteUrl to a Microsoft Graph site ID,
    2. creates a single-tenant app registration with the Sites.Selected application permission and the
       public key of -CertificatePath (no client secret),
    3. creates its service principal and grants the Sites.Selected app role (admin consent), and
    4. grants the app -Role (read or write) on that site only (POST /sites/{site-id}/permissions).

    If an app with -DisplayName already exists the script reports it and changes nothing. It never deletes
    apps. By default the script only reports. -Apply makes the changes and -WhatIf shows them.

    Keep the certificate's private key in a certificate store or Azure Key Vault, never in a file next to
    the script (docs/08). There is deliberately no -AllCustomers switch, because creating an app in every
    customer is the pattern the MspGdap partner app replaces.

.PARAMETER TenantId
    Tenant ID (GUID) or verified domain of the tenant that owns the site. Accepts pipeline input.

.PARAMETER DisplayName
    Display name of the app registration.

.PARAMETER SiteUrl
    Full URL of the site, for example https://contoso.sharepoint.com/sites/Reports.

.PARAMETER Role
    Permission on the site: read or write.

.PARAMETER CertificatePath
    Path to the public certificate (.cer) whose private key the app will sign in with.

.PARAMETER Apply
    Create the app, consent and site permission. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./new-sharepoint-site-app.ps1 -TenantId 'contoso.onmicrosoft.com' -DisplayName 'Contoso Reports Writer' -SiteUrl 'https://contoso.sharepoint.com/sites/Reports' -Role write -CertificatePath ./reports-writer.cer

    Reports what would be created, and whether the site and app name already exist.

.EXAMPLE
    ./new-sharepoint-site-app.ps1 -TenantId 'contoso.onmicrosoft.com' -DisplayName 'Contoso Reports Writer' -SiteUrl 'https://contoso.sharepoint.com/sites/Reports' -Role write -CertificatePath ./reports-writer.cer -Apply -WhatIf

    Shows each change -Apply would make.

.NOTES
    Replaces the original 2018 method: the AzureAD and AzureRM modules (New-AzureADApplication,
    New-AzureADServiceAppRoleAssignment, New-AzureRmADServicePrincipal), an app with tenant-wide
    Sites.Manage.All, a 99-year client secret exported to CSV, the v1 token endpoint, and deletion of any
    existing app whose name matched.
    Required GDAP roles: Cloud Application Administrator (app and service principal), Privileged Role
    Administrator (grant the Microsoft Graph app role), SharePoint Administrator (site permission).
    Required partner app permissions: Microsoft Graph delegated Application.ReadWrite.All,
    AppRoleAssignment.ReadWrite.All and Sites.FullControl.All (all in the full manifest).
    Microsoft Learn: https://learn.microsoft.com/en-us/graph/api/site-post-permissions

.LINK
    https://gcit.com.au/knowledge-base/create-a-sharepoint-application-for-the-microsoft-graph-via-powershell/

.LINK
    docs/02-create-partner-app.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(SupportsShouldProcess)]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [Parameter(Mandatory)]
    [ValidatePattern('^https://[A-Za-z0-9-]+\.sharepoint\.com/.+')]
    [string]$SiteUrl,

    [ValidateSet('read', 'write')]
    [string]$Role = 'read',

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$CertificatePath,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path -LiteralPath $CertificatePath).ProviderPath)
    if ($certificate.HasPrivateKey) {
        Write-Warning 'The certificate file contains a private key. Only the public key is uploaded. Keep the private key out of script folders.'
    }
    $site = [uri]$SiteUrl
    $sitePath = '{0}:{1}' -f $site.Host, $site.AbsolutePath.TrimEnd('/')

    function ConvertTo-ResultRow {
        param($Customer, $SiteId, $AppId, $ServicePrincipalId, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId      = $Customer.TenantId
            CustomerName          = $Customer.Name
            DisplayName           = $DisplayName
            SiteUrl               = $SiteUrl
            SiteId                = $SiteId
            Role                  = $Role
            AppId                 = $AppId
            ServicePrincipalId    = $ServicePrincipalId
            CertificateThumbprint = $certificate.Thumbprint
            CertificateExpires    = $certificate.NotAfter
            Action                = $Action
            Error                 = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })

    foreach ($customer in $customers) {
        $createdAppId = $null
        try {
            $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
            $customer.Name = $organisation.displayName
            if ($organisation.id) { $customer.TenantId = $organisation.id }

            $siteId = (Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "sites/$sitePath`?`$select=id").id
            if (-not $siteId) { throw "Site $SiteUrl was not found." }

            $escapedName = $DisplayName.Replace("'", "''")
            $existing = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "applications?`$filter=displayName eq '$escapedName'&`$select=id,appId")
            if ($existing.Count -gt 0) {
                $row = ConvertTo-ResultRow -Customer $customer -SiteId $siteId -AppId $existing[0].appId -Action 'AppAlreadyExists'
            }
            elseif (-not $Apply) {
                $row = ConvertTo-ResultRow -Customer $customer -SiteId $siteId -Action 'WouldCreate'
            }
            elseif ($PSCmdlet.ShouldProcess($customer.Name, "Create app '$DisplayName' with Sites.Selected and $Role access to $SiteUrl")) {
                $graphSp = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "servicePrincipals(appId='$graphAppId')?`$select=id,appRoles"
                $sitesSelected = @($graphSp.appRoles | Where-Object { $_.value -eq 'Sites.Selected' -and @($_.allowedMemberTypes) -contains 'Application' }) | Select-Object -First 1
                if (-not $sitesSelected) { throw 'The Sites.Selected app role was not found on Microsoft Graph in this tenant.' }

                $appBody = @{
                    displayName            = $DisplayName
                    signInAudience         = 'AzureADMyOrg'
                    requiredResourceAccess = @(@{
                            resourceAppId  = $graphAppId
                            resourceAccess = @(@{ id = $sitesSelected.id; type = 'Role' })
                        })
                    keyCredentials         = @(@{
                            type        = 'AsymmetricX509Cert'
                            usage       = 'Verify'
                            key         = [System.Convert]::ToBase64String($certificate.RawData)
                            displayName = "CN=$DisplayName"
                        })
                }
                $application = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri 'applications' -Body $appBody -Confirm:$false
                $createdAppId = $application.appId
                # A new application can take a few seconds to replicate before its service principal can be
                # created, so retry that step briefly.
                $servicePrincipal = $null
                for ($attempt = 1; -not $servicePrincipal; $attempt++) {
                    try {
                        $servicePrincipal = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri 'servicePrincipals' -Body @{ appId = $application.appId } -Confirm:$false
                    }
                    catch {
                        if ($attempt -ge 6) { throw }
                        Write-Verbose "Service principal not created yet (attempt $attempt): $($_.Exception.Message)"
                        Start-Sleep -Seconds 5
                    }
                }
                $grantBody = @{ principalId = $servicePrincipal.id; resourceId = $graphSp.id; appRoleId = $sitesSelected.id }
                $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri "servicePrincipals/$($graphSp.id)/appRoleAssignedTo" -Body $grantBody -Confirm:$false
                $permissionBody = @{
                    roles               = @($Role)
                    grantedToIdentities = @(@{ application = @{ id = $application.appId; displayName = $DisplayName } })
                }
                $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri "sites/$siteId/permissions" -Body $permissionBody -Confirm:$false
                $row = ConvertTo-ResultRow -Customer $customer -SiteId $siteId -AppId $application.appId -ServicePrincipalId $servicePrincipal.id -Action 'Created'
            }
            else {
                $row = ConvertTo-ResultRow -Customer $customer -SiteId $siteId -Action 'WhatIf'
            }
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Tenant $($customer.TenantId): $($_.Exception.Message)"
            # If the app was created before the failure, return its ID so it can be finished or removed.
            $failedAction = if ($createdAppId) { 'FailedAfterAppCreated' } else { 'Failed' }
            $row = ConvertTo-ResultRow -Customer $customer -AppId $createdAppId -Action $failedAction -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }
    $certificate.Dispose()

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
