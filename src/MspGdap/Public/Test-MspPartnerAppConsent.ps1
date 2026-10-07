function Test-MspPartnerAppConsent {
    <#
    .SYNOPSIS
        Checks, read-only, that the partner app is consented in a customer tenant with every expected scope.
    .DESCRIPTION
        Reads through Microsoft Graph in the customer:
          - the partner app's service principal exists (Microsoft Entra blocks multi-tenant apps without one)
          - each resource's service principal exists (Partner Center consent needs them)
          - the tenant-wide delegated grant (AllPrincipals) for each resource holds every expected scope
        Extra scopes and any app role assignments on the partner app are reported as warnings, because the
        partner app is meant to be delegated only.
        Never changes anything.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER AppId
        Partner app ID. Defaults to the configured app.
    .PARAMETER ManifestPath
        Compare against a manifest instead of the app's requiredResourceAccess. Manifest ids are checked
        against the published scope names in the partner tenant.
    .PARAMETER ExcludeResource
        Resource app IDs or display names (wildcards allowed) to leave out of the check.
    .EXAMPLE
        Test-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [string]$AppId,
        [string]$ManifestPath,
        [string[]]$ExcludeResource
    )
    begin {
        $resolvedAppId = Resolve-MspPartnerAppId -AppId $AppId
        $desiredParams = @{ ExcludeResource = $ExcludeResource; AppId = $resolvedAppId }
        if ($ManifestPath) { $desiredParams['ManifestPath'] = $ManifestPath; $desiredParams['AllowUnregisteredScope'] = $true }
        $desired = Get-MspDesiredDelegatedGrant @desiredParams
        foreach ($warning in $desired.Warnings) { Write-Warning $warning }
    }
    process {
        $steps = New-Object System.Collections.Generic.List[object]
        $info = [ordered]@{ AppId = $resolvedAppId; ServicePrincipalId = $null; MissingScopes = @{} }
        try { $resolved = Resolve-MspCustomerTenant -TenantId $TenantId }
        catch {
            $steps.Add((New-MspStepResult -Step 'Resolve tenant' -Status Failed -Detail $_.Exception.Message))
            return (New-MspOperationResult -Operation 'Test-MspPartnerAppConsent' -TenantId $TenantId -Target $resolvedAppId -Steps $steps.ToArray() -Property $info)
        }

        try { $state = Get-MspCustomerConsentState -TenantId $resolved -AppId $resolvedAppId -Grant $desired.Grants }
        catch {
            if (Test-MspConsentMissingError -ErrorRecord $_) {
                $steps.Add((New-MspStepResult -Step 'Read consent' -Status Failed -Detail 'Not consented (AADSTS65001). Run Grant-MspPartnerAppConsent for this customer.'))
                return (New-MspOperationResult -Operation 'Test-MspPartnerAppConsent' -TenantId $resolved -Target $resolvedAppId -Steps $steps.ToArray() -Property $info)
            }
            $steps.Add((New-MspStepResult -Step 'Read consent' -Status Unknown -Detail "Could not read the customer through Graph. $($_.Exception.Message)"))
            return (New-MspOperationResult -Operation 'Test-MspPartnerAppConsent' -TenantId $resolved -Target $resolvedAppId -Steps $steps.ToArray() -Property $info)
        }

        if ($state.AppServicePrincipal) {
            $info.ServicePrincipalId = [string]$state.AppServicePrincipal.id
            $steps.Add((New-MspStepResult -Step 'Partner app service principal' -Status Passed -Detail "Present ($($state.AppServicePrincipal.id))."))
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Partner app service principal' -Status Failed -Detail 'Not present. Run Grant-MspPartnerAppConsent for this customer.'))
        }

        foreach ($resource in $state.Resources) {
            $name = "Consent: $($resource.ResourceDisplayName)"
            if (-not $resource.ResourcePresent) {
                $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail 'Resource service principal is not present in this customer.'))
                $info.MissingScopes[$resource.ResourceDisplayName] = $resource.DesiredScopes
                continue
            }
            if (-not $resource.GrantId) {
                $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "No tenant-wide consent. Missing: $($resource.DesiredScopes -join ', ')"))
                $info.MissingScopes[$resource.ResourceDisplayName] = $resource.DesiredScopes
                continue
            }
            if ($resource.MissingScopes.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "Missing: $($resource.MissingScopes -join ', ')"))
                $info.MissingScopes[$resource.ResourceDisplayName] = $resource.MissingScopes
            }
            else {
                $steps.Add((New-MspStepResult -Step $name -Status Passed -Detail "All $($resource.DesiredScopes.Count) scopes present."))
            }
            if ($resource.ExtraScopes.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step "Extra scopes: $($resource.ResourceDisplayName)" -Status Warning -Detail "Consented but not expected: $($resource.ExtraScopes -join ', ')"))
            }
        }

        if (@($state.AppRoleAssignments).Count -gt 0) {
            $steps.Add((New-MspStepResult -Step 'Application permissions on partner app' -Status Warning -Detail "The partner app holds $(@($state.AppRoleAssignments).Count) app role assignments in this customer. MspGdap's partner app is meant to be delegated only. Review them." -Data $state.AppRoleAssignments))
        }

        New-MspOperationResult -Operation 'Test-MspPartnerAppConsent' -TenantId $resolved -Target $resolvedAppId -Steps $steps.ToArray() -Property $info
    }
}
