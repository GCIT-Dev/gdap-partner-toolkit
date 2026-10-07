function Grant-MspPartnerAppConsent {
    <#
    .SYNOPSIS
        Pre-consents the partner app's delegated scopes into GDAP customer tenants through Partner Center.
    .DESCRIPTION
        Uses Partner Center POST /v1/customers/{customerId}/applicationconsents, as Microsoft documents for
        GDAP and the Secure Application Model (GDAP FAQ):
          - one resource per POST, scopes joined with ", "
          - App+User call with the technician's MFA-backed token, ValidateMfa: true header, and an
            ms-requestid idempotency key that is reused if the call is retried
          - the token is always minted for the configured partner app, and Microsoft requires the token's
            app ID to match applicationId in the body, so -AppId must be the configured app
          - the technician needs Application Administrator or Cloud Application Administrator in the customer
            through GDAP, and membership of AdminAgents in the partner tenant
        Idempotent: the current consent is read through Microsoft Graph in the customer first, and only
        resources with no consent are posted. Resources whose consent exists but lacks scopes need -Force,
        because Microsoft's documented way to add scopes is DELETE then POST of the app consent.
        With -Force every scope the app holds in that customer now (including resources you excluded or
        that are not in the manifest) is posted again together with the missing scopes, so nothing that
        was consented before is dropped. The confirmation message lists exactly what will be re-posted.
        Every change is read back through Graph before it is reported as Changed. A Failed result is also
        written as a non-terminating error (MspGdap.Grant-MspPartnerAppConsent.Failed).
        Application (app-only) permissions cannot be consented this way under GDAP and are skipped.
        Your partner tenant is refused as a target.
    .PARAMETER TenantId
        One or more customer tenant IDs or verified domains. Alias: CustomerTenantId.
    .PARAMETER AllCustomers
        Every customer returned by Get-MspCustomer.
    .PARAMETER AppId
        Partner app ID. Defaults to the configured app (Set-MspConfiguration). Any other value is refused,
        because the Partner Center token is always minted for the configured app and Microsoft rejects a
        mismatch.
    .PARAMETER ManifestPath
        Take the desired scopes from a manifest instead of the app's requiredResourceAccess. Each permission
        id is checked against its value through the resource's published scopes in the partner tenant, and
        a scope the partner app registration does not request is refused unless -Force is given.
    .PARAMETER DisplayName
        Display name sent to Partner Center. Defaults to the partner app's display name.
    .PARAMETER ExcludeResource
        Resource app IDs or display names (wildcards allowed) to leave out, for example 'WindowsDefenderATP*'.
    .PARAMETER Force
        Allow DELETE then POST of the app consent when an existing consent is missing scopes. The app has no
        consent in that customer for a few seconds while this runs. Also allows manifest scopes that the
        partner app registration does not request.
    .PARAMETER SkipGdapCheck
        Do not check for an active GDAP relationship first.
    .PARAMETER ReadbackTimeoutSeconds
        How long to keep reading the consent back through Graph before reporting it. Default 60.
    .EXAMPLE
        Grant-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com' -WhatIf
    .EXAMPLE
        Get-MspCustomer | Grant-MspPartnerAppConsent -ExcludeResource 'WindowsDefenderATP*'
    .OUTPUTS
        One MspGdap.OperationResult per customer.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string[]]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'All')][switch]$AllCustomers,
        [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [string]$ManifestPath,
        [string]$DisplayName,
        [string[]]$ExcludeResource,
        [switch]$Force,
        [switch]$SkipGdapCheck,
        [ValidateRange(0, 600)][int]$ReadbackTimeoutSeconds = 60
    )
    begin {
        try {
            $resolvedAppId = Resolve-MspPartnerAppId
            if ($AppId -and $AppId.ToLowerInvariant() -ne $resolvedAppId) {
                throw (New-MspErrorRecord -Message "AppId $AppId is not the configured partner app ($resolvedAppId). Partner Center tokens are always issued to the configured app, and Microsoft rejects a consent whose applicationId differs from the token. Run Set-MspConfiguration -AppId first if you mean to switch apps." -ErrorId 'MspGdap.Consent.AppIdMismatch' -Category InvalidArgument -TargetObject $AppId)
            }
            $desiredParams = @{ ExcludeResource = $ExcludeResource; AppId = $resolvedAppId }
            if ($ManifestPath) {
                $desiredParams['ManifestPath'] = $ManifestPath
                $desiredParams['AllowUnregisteredScope'] = [bool]$Force
            }
            $desired = Get-MspDesiredDelegatedGrant @desiredParams
            foreach ($warning in $desired.Warnings) { Write-Warning $warning }
            if (@($desired.Grants).Count -eq 0) {
                throw (New-MspErrorRecord -Message 'No delegated scopes to consent. Check the app or manifest and -ExcludeResource.' -ErrorId 'MspGdap.Consent.NothingToConsent' -Category InvalidData -TargetObject $resolvedAppId)
            }
            if ($desired.SkippedRoles.Count -gt 0) {
                Write-Verbose ("Application permissions are not consented through Partner Center under GDAP and were skipped: {0}" -f ($desired.SkippedRoles -join ', '))
            }

            $appDisplayName = $DisplayName
            if (-not $appDisplayName) { $appDisplayName = $desired.AppDisplayName }
            if (-not $appDisplayName) {
                try {
                    $partnerApp = @(Invoke-MspGraphCall -PartnerTenant -Path ("applications?`$filter=appId eq '{0}'&`$select=displayName" -f $resolvedAppId)) | Select-Object -First 1
                    if ($partnerApp) { $appDisplayName = [string]$partnerApp.displayName }
                }
                catch { Write-Verbose "Could not read the partner app registration: $($_.Exception.Message)" }
            }
            if (-not $appDisplayName) {
                $partnerSp = Get-MspServicePrincipalByAppId -PartnerTenant -AppId $resolvedAppId
                if ($partnerSp) { $appDisplayName = [string]$partnerSp.displayName }
            }
            if (-not $appDisplayName) {
                throw (New-MspErrorRecord -Message "Could not read the display name of app $resolvedAppId in the partner tenant. Pass -DisplayName." -ErrorId 'MspGdap.Consent.NoDisplayName' -Category ObjectNotFound -TargetObject $resolvedAppId)
            }

            $activeByTenant = @{}
            if (-not $SkipGdapCheck) {
                foreach ($relationship in @(Get-MspGdapRelationship -Status active)) {
                    $key = ([string]$relationship.CustomerTenantId).ToLowerInvariant()
                    if (-not $activeByTenant.ContainsKey($key)) { $activeByTenant[$key] = New-Object System.Collections.Generic.List[object] }
                    $activeByTenant[$key].Add($relationship)
                }
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        $targets = New-Object System.Collections.Generic.List[object]
        $cache = @{ Customers = $null }
    }
    process {
        if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
            foreach ($id in $TenantId) { $targets.Add([pscustomobject]@{ TenantId = $id; DisplayName = $null; PartnerCenterId = $null }) }
        }
    }
    end {
        if ($AllCustomers) {
            foreach ($customer in @(Get-MspCustomer | ConvertTo-MspCustomerReference)) { if ($customer.TenantId) { $targets.Add($customer) } }
        }
        $stopAll = $false

        foreach ($target in $targets) {
            $steps = New-Object System.Collections.Generic.List[object]
            $info = [ordered]@{ AppId = $resolvedAppId; CustomerName = $target.DisplayName; IsMfaCompliant = $null }
            $customerTenant = $null
            try { $customerTenant = Resolve-MspCustomerTenant -TenantId $target.TenantId }
            catch {
                $steps.Add((New-MspStepResult -Step 'Resolve tenant' -Status Failed -Detail $_.Exception.Message))
                New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Grant-MspPartnerAppConsent' -TenantId $target.TenantId -Target $appDisplayName -Steps $steps.ToArray() -Property $info
                continue
            }
            $emit = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Grant-MspPartnerAppConsent' -TenantId $customerTenant -Target $appDisplayName -Steps $steps.ToArray() -Property $info }

            if ($stopAll) {
                $steps.Add((New-MspStepResult -Step 'Partner Center MFA' -Status Failed -Detail 'Skipped: an earlier Partner Center call was rejected with MFA required.'))
                & $emit; continue
            }

            if (-not $SkipGdapCheck) {
                $active = if ($activeByTenant.ContainsKey($customerTenant)) { $activeByTenant[$customerTenant] } else { @() }
                if (@($active).Count -eq 0) {
                    $steps.Add((New-MspStepResult -Step 'Check GDAP relationship' -Status Failed -Detail 'No active GDAP relationship with this customer. Create and get approval first (New-MspGdapRelationship).'))
                    & $emit; continue
                }
                if (-not $info.CustomerName) { $info.CustomerName = @($active)[0].CustomerName }
                $steps.Add((New-MspStepResult -Step 'Check GDAP relationship' -Status Passed -Detail ("{0} active: {1}" -f @($active).Count, ((@($active) | ForEach-Object { $_.DisplayName }) -join ', '))))
            }

            # Current state through Graph
            # The read uses the partner app's own token in the customer, so before the first consent it fails with
            # AADSTS65001. That means "nothing consented yet", not "unknown": post every resource and confirm afterwards.
            $notConsentedYet = $false
            try { $state = Get-MspCustomerConsentState -TenantId $customerTenant -AppId $resolvedAppId -Grant $desired.Grants }
            catch {
                if (Test-MspConsentMissingError -ErrorRecord $_) {
                    $notConsentedYet = $true
                    $state = Get-MspEmptyConsentState -TenantId $customerTenant -Grant $desired.Grants
                }
                else {
                    $steps.Add((New-MspStepResult -Step 'Read current consent' -Status Unknown -Detail "Could not read the customer through Graph, so the current consent is unknown and nothing was posted. $($_.Exception.Message)"))
                    & $emit; continue
                }
            }
            if ($notConsentedYet) {
                $steps.Add((New-MspStepResult -Step 'Read current consent' -Status Passed -Detail 'Not consented yet (AADSTS65001). The app cannot read this customer until consent exists, so every resource will be posted and then confirmed through Graph.'))
            }
            else {
                $appSpText = if ($state.AppServicePrincipal) { "service principal $($state.AppServicePrincipal.id)" } else { 'no service principal yet' }
                $steps.Add((New-MspStepResult -Step 'Read current consent' -Status Passed -Detail "Partner app has $appSpText."))
            }

            $toCreate = New-Object System.Collections.Generic.List[object]
            $toTopUp = New-Object System.Collections.Generic.List[object]
            foreach ($resource in $state.Resources) {
                $name = "Consent: $($resource.ResourceDisplayName)"
                if (-not $resource.ResourcePresent) {
                    $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail 'The resource has no service principal in this customer (often an unlicensed workload). Partner Center cannot consent to it. Exclude it with -ExcludeResource or license the workload.'))
                }
                elseif ($resource.MissingScopes.Count -eq 0) {
                    $steps.Add((New-MspStepResult -Step $name -Status Passed -Detail "All $($resource.DesiredScopes.Count) scopes already consented."))
                }
                elseif (-not $resource.GrantId) { $toCreate.Add($resource) }
                else { $toTopUp.Add($resource) }
            }

            $posted = New-Object System.Collections.Generic.List[object]
            $checkGrants = @($desired.Grants)
            # Partner Center documents the path ID as "ID of the customer generated in Partner Center". It is
            # normally the tenant ID. On a 404 the Partner Center customer list is checked for a different ID.
            $pc = @{ CustomerId = $customerTenant }
            $postConsent = {
                param($Resource)
                $body = @{
                    applicationId     = $resolvedAppId
                    displayName       = $appDisplayName
                    applicationGrants = @(@{ enterpriseApplicationId = $Resource.ResourceAppId; scope = ($Resource.DesiredScopes -join ', ') })
                }
                $response = Invoke-MspPartnerCenterRequest -Method POST -Path "customers/$($pc.CustomerId)/applicationconsents" -Body $body
                if ($response.StatusCode -eq 404 -and $pc.CustomerId -eq $customerTenant) {
                    if ($null -eq $cache.Customers) { $cache.Customers = @(Get-MspCustomer | ConvertTo-MspCustomerReference) }
                    $match = @($cache.Customers | Where-Object { $_.TenantId -eq $customerTenant -and $_.PartnerCenterId -and $_.PartnerCenterId -ne $customerTenant }) | Select-Object -First 1
                    if ($match) {
                        $pc.CustomerId = $match.PartnerCenterId
                        $response = Invoke-MspPartnerCenterRequest -Method POST -Path "customers/$($pc.CustomerId)/applicationconsents" -Body $body
                    }
                }
                $response
            }

            $mfaFailed = $false
            foreach ($resource in $toCreate) {
                $name = "Consent: $($resource.ResourceDisplayName)"
                if (-not $PSCmdlet.ShouldProcess("customer $customerTenant", "Pre-consent '$appDisplayName' for $($resource.ResourceDisplayName): $($resource.DesiredScopes -join ', ')")) {
                    $steps.Add((New-MspStepResult -Step $name -Status WhatIf -Detail "Would POST $($resource.DesiredScopes.Count) scopes."))
                    continue
                }
                $response = & $postConsent $resource
                if ($null -ne $response.IsMfaCompliant) { $info.IsMfaCompliant = $response.IsMfaCompliant }
                if ($response.MfaRequired) {
                    $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail $response.Error)); $mfaFailed = $true; break
                }
                switch ($response.StatusCode) {
                    { $_ -in 200, 201 } { $posted.Add($resource); break }
                    409 { $posted.Add($resource); Write-Verbose "Partner Center reports a consent already exists for $($resource.ResourceDisplayName). Confirming through Graph."; break }
                    403 { $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "Partner Center returned 403. Check that you are in AdminAgents and hold Cloud Application Administrator or Application Administrator in this customer through GDAP. $($response.Error)")); break }
                    default { $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "Partner Center returned $($response.StatusCode). $($response.Error)")) }
                }
            }

            if (-not $mfaFailed -and $toTopUp.Count -gt 0) {
                $missingText = ($toTopUp | ForEach-Object { "$($_.ResourceDisplayName): $($_.MissingScopes -join ', ')" }) -join '; '
                if (-not $Force) {
                    foreach ($resource in $toTopUp) {
                        $steps.Add((New-MspStepResult -Step "Consent: $($resource.ResourceDisplayName)" -Status Failed -Detail "Existing consent is missing: $($resource.MissingScopes -join ', '). Re-run with -Force to replace the consent (Partner Center DELETE then POST)."))
                    }
                }
                else {
                    # Build the complete re-post list first: everything consented now plus what is missing.
                    $repost = New-Object System.Collections.Generic.List[object]
                    $repostProblem = $null
                    foreach ($resource in @($state.Resources | Where-Object { $_.ResourcePresent })) {
                        $union = @(@($resource.CurrentScopes) + @($resource.DesiredScopes) | Where-Object { $_ } | Select-Object -Unique)
                        $repost.Add([pscustomobject]@{ ResourceAppId = $resource.ResourceAppId; ResourceDisplayName = $resource.ResourceDisplayName; DesiredScopes = $union; Scopes = $union })
                    }
                    foreach ($other in @($state.OtherGrants)) {
                        try {
                            $otherSp = Invoke-MspGraphCall -TenantId $customerTenant -Path ("servicePrincipals/{0}?`$select=id,appId,displayName" -f $other.ResourceSpId)
                            if (-not $otherSp -or -not $otherSp.appId) { throw "Service principal $($other.ResourceSpId) was not found." }
                            $repost.Add([pscustomobject]@{ ResourceAppId = [string]$otherSp.appId; ResourceDisplayName = [string]$otherSp.displayName; DesiredScopes = @($other.Scopes); Scopes = @($other.Scopes) })
                        }
                        catch {
                            $repostProblem = "Could not identify the resource of an existing grant ($($other.ResourceSpId)), so replacing the consent could drop it. Nothing was deleted. $($_.Exception.Message)"
                        }
                    }
                    $plan = ($repost | ForEach-Object { "$($_.ResourceDisplayName): $($_.DesiredScopes -join ', ')" }) -join '; '
                    if ($repostProblem) {
                        $steps.Add((New-MspStepResult -Step 'Replace consent' -Status Failed -Detail $repostProblem))
                    }
                    elseif ($PSCmdlet.ShouldProcess("customer $customerTenant", "Replace the whole consent for '$appDisplayName' (DELETE, then POST each resource) to add $missingText. Re-posted after the DELETE: $plan. Nothing consented now is dropped")) {
                        $delete = Invoke-MspPartnerCenterRequest -Method DELETE -Path "customers/$($pc.CustomerId)/applicationconsents/$resolvedAppId"
                        if ($null -ne $delete.IsMfaCompliant) { $info.IsMfaCompliant = $delete.IsMfaCompliant }
                        if ($delete.MfaRequired) {
                            $steps.Add((New-MspStepResult -Step 'Replace consent' -Status Failed -Detail $delete.Error)); $mfaFailed = $true
                        }
                        elseif (-not $delete.Success -and $delete.StatusCode -ne 404) {
                            $steps.Add((New-MspStepResult -Step 'Replace consent' -Status Failed -Detail "DELETE returned $($delete.StatusCode). The existing consent was left in place. $($delete.Error)"))
                        }
                        else {
                            $steps.Add((New-MspStepResult -Step 'Replace consent' -Status Passed -Detail "Existing consent removed. Posting $($repost.Count) resources again."))
                            $checkGrants = @($repost | ForEach-Object { [pscustomobject]@{ ResourceAppId = $_.ResourceAppId; ResourceDisplayName = $_.ResourceDisplayName; Scopes = $_.Scopes } })
                            foreach ($resource in $repost) {
                                $response = & $postConsent $resource
                                if ($response.StatusCode -in 200, 201, 409) {
                                    $already = @($posted | Where-Object { $_.ResourceAppId -eq $resource.ResourceAppId })
                                    if ($already.Count -eq 0) { $posted.Add($resource) }
                                }
                                else { $steps.Add((New-MspStepResult -Step "Consent: $($resource.ResourceDisplayName)" -Status Failed -Detail "Re-POST after DELETE returned $($response.StatusCode). The app may now be missing this resource. Scopes to restore: $($resource.DesiredScopes -join ', '). $($response.Error)")) }
                            }
                        }
                    }
                    else {
                        foreach ($resource in $toTopUp) {
                            $steps.Add((New-MspStepResult -Step "Consent: $($resource.ResourceDisplayName)" -Status WhatIf -Detail "Would replace the consent to add: $($resource.MissingScopes -join ', '). Would re-post: $plan"))
                        }
                    }
                }
            }
            if ($mfaFailed) { $stopAll = $true }

            if ($posted.Count -gt 0) {
                $postedIds = @($posted | ForEach-Object { $_.ResourceAppId })
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                    $now = Get-MspCustomerConsentState -TenantId $customerTenant -AppId $resolvedAppId -Grant $checkGrants
                    $pending = @($now.Resources | Where-Object { $postedIds -contains $_.ResourceAppId -and $_.MissingScopes.Count -gt 0 })
                    if ($pending.Count -eq 0) { $now }
                }
                $final = $check.Value
                if (-not $final) {
                    try { $final = Get-MspCustomerConsentState -TenantId $customerTenant -AppId $resolvedAppId -Grant $checkGrants } catch { $final = $null }
                }
                foreach ($resource in $posted) {
                    $name = "Consent: $($resource.ResourceDisplayName)"
                    $now = if ($final) { @($final.Resources | Where-Object { $_.ResourceAppId -eq $resource.ResourceAppId }) | Select-Object -First 1 } else { $null }
                    if (-not $now) {
                        $steps.Add((New-MspStepResult -Step $name -Status Unknown -Detail "Partner Center accepted the request, but Graph could not be read to confirm it. $($check.LastError)"))
                    }
                    elseif ($now.MissingScopes.Count -eq 0) {
                        $steps.Add((New-MspStepResult -Step $name -Status Changed -Detail "Consented and confirmed through Graph: $($now.DesiredScopes -join ', ')"))
                    }
                    else {
                        $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "Partner Center accepted the request but Graph still shows missing: $($now.MissingScopes -join ', '). If an older consent exists, re-run with -Force."))
                    }
                }
            }
            & $emit
        }
    }
}
