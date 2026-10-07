function New-MspPartnerApp {
    <#
    .SYNOPSIS
        Creates the multi-tenant partner app in YOUR partner tenant from a permission manifest.
    .DESCRIPTION
        Run once per partner. The command:
          1. checks it is connected to the partner tenant you name (never any other tenant)
          2. loads the manifest and checks every permission ID against the resource's published permissions
          3. creates the application (multi-tenant, Web redirect http://localhost, no implicit grant,
             app instance property lock on, certificate public key if supplied)
          4. creates its service principal in the partner tenant (Microsoft Entra blocks multi-tenant apps
             without a service principal from 31 March 2026)
          5. grants tenant-wide admin consent for the delegated scopes in the partner tenant
          6. optionally restricts sign-in to one security group of technicians
        Every write is read back. Steps that cannot be confirmed are reported as Failed, and a Failed result is
        also written as a non-terminating error (MspGdap.New-MspPartnerApp.Failed).
        -RestrictToGroupId is recommended: only your technicians' group can then sign in to the app.

        Before the app exists, MspGdap has no token of its own, so the default transport is the Microsoft Graph
        PowerShell SDK. Connect first with an account that can create applications in the partner tenant:
          Connect-MgGraph -TenantId <PartnerTenantId> -Scopes Application.ReadWrite.All,DelegatedPermissionGrant.ReadWrite.All,AppRoleAssignment.ReadWrite.All

        Application (Role) permissions in the manifest are written to the app but never granted here. Grant them
        per customer with Enable-MspExchangeAppAccess (automation app only).
    .PARAMETER DisplayName
        Name of the new application, for example 'Contoso MSP Worker'. Customers see it under Enterprise applications.
    .PARAMETER PartnerTenantId
        Tenant ID (GUID) of your partner tenant. The command stops if the Graph session is in any other tenant.
    .PARAMETER ManifestPath
        Permission manifest. Defaults to manifests/partner-app.minimal.json.
    .PARAMETER CertificateThumbprint
        Thumbprint of a certificate in CurrentUser\My or LocalMachine\My. Only the public key is uploaded.
    .PARAMETER Certificate
        An X509Certificate2 object. Only the public key is uploaded.
    .PARAMETER CertificatePath
        Path to a public certificate file (.cer, .crt or .pem). PFX files are refused.
    .PARAMETER RedirectUri
        Web platform redirect URIs. Default http://localhost (the port is ignored by Microsoft Entra ID for localhost).
    .PARAMETER RestrictToGroupId
        Object ID of a partner tenant security group. When set, user assignment is required on the partner
        tenant service principal and only this group can sign in to the app there.
    .PARAMETER SkipAdminConsent
        Do not grant admin consent in the partner tenant.
    .PARAMETER UseExisting
        If an application with this display name already exists, reuse it instead of stopping. Permissions
        in the manifest are ADDED to the app's existing permissions and missing redirect URIs are added.
        Nothing the app already has is removed unless -RemoveUnlisted is also given.
    .PARAMETER RemoveUnlisted
        With -UseExisting: also remove permissions that are not in the manifest, so the app matches it exactly.
    .PARAMETER Transport
        MgGraph (default, Microsoft.Graph.Authentication session) or MspGdap (an existing MspGdap partner tenant token).
    .PARAMETER ReadbackTimeoutSeconds
        How long to keep reading each change back before reporting it as Failed. Default 60.
    .EXAMPLE
        Connect-MgGraph -TenantId '<PartnerTenantId>' -Scopes Application.ReadWrite.All,DelegatedPermissionGrant.ReadWrite.All,AppRoleAssignment.ReadWrite.All
        New-MspPartnerApp -DisplayName 'Contoso MSP Worker' -PartnerTenantId '<PartnerTenantId>' -CertificateThumbprint '<Thumbprint>' -WhatIf
    .OUTPUTS
        MspGdap.OperationResult with AppId, ApplicationObjectId, ServicePrincipalId and NextSteps.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'NoCertificate')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][ValidateLength(1, 120)][string]$DisplayName,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$PartnerTenantId,
        [string]$ManifestPath,
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$CertificateThumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory, ParameterSetName = 'Path')][string]$CertificatePath,
        [ValidateNotNullOrEmpty()][string[]]$RedirectUri = @('http://localhost'),
        [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$RestrictToGroupId,
        [switch]$SkipAdminConsent,
        [switch]$UseExisting,
        [switch]$RemoveUnlisted,
        [ValidateSet('MgGraph', 'MspGdap')][string]$Transport = 'MgGraph',
        [ValidateRange(0, 600)][int]$ReadbackTimeoutSeconds = 60
    )

    $PartnerTenantId = $PartnerTenantId.ToLowerInvariant()
    $steps = New-Object System.Collections.Generic.List[object]
    $state = [ordered]@{ AppId = $null; ApplicationObjectId = $null; ServicePrincipalId = $null; CertificateThumbprint = $null; NextSteps = @() }
    $finish = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'New-MspPartnerApp' -TenantId $PartnerTenantId -Target $DisplayName -Steps $steps.ToArray() -Property $state }
    $g = @{ Transport = $Transport }

    # 1. Connection
    $contextStep = Test-MspSetupContext -Transport $Transport -PartnerTenantId $PartnerTenantId
    $steps.Add($contextStep)
    if ($contextStep.Status -eq 'Failed') { return (& $finish) }

    # 2. Certificate (public key only)
    $cert = $null
    if ($PSCmdlet.ParameterSetName -ne 'NoCertificate') {
        try {
            $cert = switch ($PSCmdlet.ParameterSetName) {
                'Thumbprint' { Resolve-MspCertificate -Thumbprint $CertificateThumbprint }
                'Certificate' { Resolve-MspCertificate -Certificate $Certificate }
                'Path' { Resolve-MspCertificate -Path $CertificatePath }
            }
            $state.CertificateThumbprint = $cert.Thumbprint
            $steps.Add((New-MspStepResult -Step 'Check certificate' -Status Passed -Detail ("{0}, expires {1:yyyy-MM-dd}. Only the public key is uploaded." -f $cert.Thumbprint, $cert.NotAfter)))
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Check certificate' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }
    }
    else {
        $steps.Add((New-MspStepResult -Step 'Check certificate' -Status Warning -Detail 'No certificate supplied. The app cannot redeem tokens until you add one with Add-MspPartnerAppCertificate.'))
    }

    # 3. Manifest
    try {
        if (-not $ManifestPath) { $ManifestPath = Resolve-MspDataPath -FileName 'partner-app.minimal.json' }
        $manifest = Read-MspPermissionManifest -Path $ManifestPath
        $steps.Add((New-MspStepResult -Step 'Load manifest' -Status Passed -Detail ("{0}: {1} resources." -f $manifest.Name, @($manifest.RequiredResourceAccess).Count)))
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Load manifest' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }

    # 4. Validate every permission against the resource service principal in the partner tenant
    $resourceSps = @{}
    $validationFailed = $false
    foreach ($resource in $manifest.RequiredResourceAccess) {
        $stepName = "Validate permissions: $($resource.resourceDisplayName)"
        try {
            $path = "servicePrincipals?`$filter=appId eq '{0}'&`$select=id,appId,displayName,oauth2PermissionScopes,appRoles" -f $resource.resourceAppId
            $sp = @(Invoke-MspPartnerGraph @g -Path $path) | Select-Object -First 1
        }
        catch {
            $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail $_.Exception.Message)); $validationFailed = $true; continue
        }
        if (-not $sp) {
            $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail "Resource $($resource.resourceAppId) has no service principal in the partner tenant, so its permissions cannot be checked or consented.")); $validationFailed = $true; continue
        }
        $resourceSps[$resource.resourceAppId] = $sp
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $resource.resourceAccess) {
            $published = if ($entry.type -eq 'Scope') { @($sp.oauth2PermissionScopes) } else { @($sp.appRoles) }
            $match = @($published | Where-Object { $_.id -eq $entry.id }) | Select-Object -First 1
            if (-not $match) { $problems.Add("$($entry.type) $($entry.id) ($($entry.value)) is not published by $($sp.displayName)"); continue }
            if ($entry.value -and ([string]$match.value -cne $entry.value)) { $problems.Add("$($entry.id) is '$($match.value)', manifest says '$($entry.value)'") }
        }
        if ($problems.Count -gt 0) {
            $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail ($problems -join '; '))); $validationFailed = $true
        }
        else {
            $steps.Add((New-MspStepResult -Step $stepName -Status Passed -Detail ("{0} permissions match." -f @($resource.resourceAccess).Count)))
        }
    }
    if ($validationFailed) { return (& $finish) }

    # 5. Existing application with the same name
    $existing = @()
    try {
        $filterName = [uri]::EscapeDataString($DisplayName.Replace("'", "''"))
        $existing = @(Invoke-MspPartnerGraph @g -Path ("applications?`$filter=displayName eq '{0}'&`$select=id,appId,displayName,signInAudience,web,requiredResourceAccess,keyCredentials" -f $filterName))
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Find existing application' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }
    $application = $null
    if ($existing.Count -gt 1) {
        $steps.Add((New-MspStepResult -Step 'Find existing application' -Status Failed -Detail "$($existing.Count) applications are named '$DisplayName'. Rename or remove the extras, then run again."))
        return (& $finish)
    }
    elseif ($existing.Count -eq 1 -and -not $UseExisting) {
        $steps.Add((New-MspStepResult -Step 'Find existing application' -Status Failed -Detail "An application named '$DisplayName' already exists (appId $($existing[0].appId)). Use -UseExisting to reuse it, or choose another name."))
        return (& $finish)
    }
    elseif ($existing.Count -eq 1) {
        $application = $existing[0]
        $steps.Add((New-MspStepResult -Step 'Find existing application' -Status Passed -Detail "Reusing application $($application.appId)."))
    }
    else {
        $steps.Add((New-MspStepResult -Step 'Find existing application' -Status Passed -Detail 'No application with this name. A new one will be created.'))
    }

    $desiredRra = $manifest.GraphRequiredResourceAccess
    $keyCredential = if ($cert) { ConvertTo-MspKeyCredential -Certificate $cert } else { $null }

    # 6. Create or update the application
    if (-not $application) {
        $body = [ordered]@{
            displayName                       = $DisplayName
            signInAudience                    = 'AzureADMultipleOrgs'
            isFallbackPublicClient            = $false
            requiredResourceAccess            = $desiredRra
            web                               = @{ redirectUris = @($RedirectUri); implicitGrantSettings = @{ enableIdTokenIssuance = $false; enableAccessTokenIssuance = $false } }
            servicePrincipalLockConfiguration = @{ isEnabled = $true; allProperties = $true }
        }
        if ($keyCredential) { $body['keyCredentials'] = @($keyCredential) }
        if ($PSCmdlet.ShouldProcess("partner tenant $PartnerTenantId", "Create multi-tenant application '$DisplayName' with $(@($desiredRra).Count) resources from $($manifest.Name)")) {
            try {
                $application = Invoke-MspPartnerGraph @g -Method POST -Path 'applications' -Body $body
                $steps.Add((New-MspStepResult -Step 'Create application' -Status Changed -Detail "Created appId $($application.appId)."))
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'Create application' -Status Failed -Detail $_.Exception.Message))
                return (& $finish)
            }
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Create application' -Status WhatIf -Detail "Would create '$DisplayName' (multi-tenant, redirect $($RedirectUri -join ', '), $(@($desiredRra).Count) resources)."))
            foreach ($later in @('Create service principal', 'Grant admin consent in partner tenant')) {
                $steps.Add((New-MspStepResult -Step $later -Status WhatIf -Detail 'Runs after the application is created.'))
            }
            return (& $finish)
        }
    }
    else {
        $patch = [ordered]@{}
        $currentRedirects = @($application.web.redirectUris | Where-Object { $_ })
        $missingRedirects = @($RedirectUri | Where-Object { $currentRedirects -notcontains $_ })
        if ($missingRedirects.Count -gt 0) {
            # PATCH replaces the whole web object, so send every existing web property with the new redirect URIs.
            $web = [ordered]@{}
            if ($application.web) {
                foreach ($property in $application.web.PSObject.Properties) { $web[$property.Name] = $property.Value }
            }
            $web['redirectUris'] = @($currentRedirects + $missingRedirects | Select-Object -Unique)
            $patch['web'] = $web
        }
        # Merge: keep what the app already has, add what the manifest asks for. -RemoveUnlisted makes it exact.
        $merged = [ordered]@{}
        $sources = if ($RemoveUnlisted) { @($desiredRra) } else { @($application.requiredResourceAccess) + @($desiredRra) }
        foreach ($r in $sources) {
            if (-not $r) { continue }
            $key = ([string]$r.resourceAppId).ToLowerInvariant()
            if (-not $merged.Contains($key)) { $merged[$key] = [ordered]@{ resourceAppId = [string]$r.resourceAppId; access = [ordered]@{} } }
            foreach ($a in @($r.resourceAccess)) {
                $accessKey = ('{0}|{1}' -f $a.id, $a.type).ToLowerInvariant()
                if (-not $merged[$key].access.Contains($accessKey)) { $merged[$key].access[$accessKey] = [pscustomobject]@{ id = [string]$a.id; type = [string]$a.type } }
            }
        }
        $mergedRra = @(foreach ($entry in $merged.Values) { [pscustomobject]@{ resourceAppId = $entry.resourceAppId; resourceAccess = @($entry.access.Values) } })
        $currentPairs = @(foreach ($r in @($application.requiredResourceAccess)) { foreach ($a in @($r.resourceAccess)) { ('{0}|{1}|{2}' -f $r.resourceAppId, $a.id, $a.type).ToLowerInvariant() } })
        $mergedPairs = @(foreach ($r in $mergedRra) { foreach ($a in @($r.resourceAccess)) { ('{0}|{1}|{2}' -f $r.resourceAppId, $a.id, $a.type).ToLowerInvariant() } })
        $added = @($mergedPairs | Where-Object { $currentPairs -notcontains $_ })
        $removed = @($currentPairs | Where-Object { $mergedPairs -notcontains $_ })
        if ($added.Count -gt 0 -or $removed.Count -gt 0) { $patch['requiredResourceAccess'] = $mergedRra }
        if ($removed.Count -gt 0) {
            $steps.Add((New-MspStepResult -Step 'Permissions not in manifest' -Status Warning -Detail "-RemoveUnlisted will remove $($removed.Count) permission(s) the manifest does not list."))
        }
        elseif (-not $RemoveUnlisted) {
            $extra = @($currentPairs | Where-Object { @(foreach ($r in @($desiredRra)) { foreach ($a in @($r.resourceAccess)) { ('{0}|{1}|{2}' -f $r.resourceAppId, $a.id, $a.type).ToLowerInvariant() } }) -notcontains $_ })
            if ($extra.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step 'Permissions not in manifest' -Status Warning -Detail "The app keeps $($extra.Count) permission(s) the manifest does not list. Use -RemoveUnlisted to remove them."))
            }
        }
        if ($patch.Count -eq 0) {
            $steps.Add((New-MspStepResult -Step 'Update application' -Status Passed -Detail 'Redirect URIs and permissions already match the manifest.'))
        }
        elseif ($PSCmdlet.ShouldProcess("application $($application.appId)", "Update $(@($patch.Keys) -join ' and ') from $($manifest.Name) (permissions: $($added.Count) added, $($removed.Count) removed)")) {
            try {
                $null = Invoke-MspPartnerGraph @g -Method PATCH -Path "applications/$($application.id)" -Body $patch
                $steps.Add((New-MspStepResult -Step 'Update application' -Status Changed -Detail "Updated $(@($patch.Keys) -join ', ')."))
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'Update application' -Status Failed -Detail $_.Exception.Message))
                return (& $finish)
            }
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Update application' -Status WhatIf -Detail "Would update $(@($patch.Keys) -join ', ')."))
        }
        if ($cert -and -not (Test-MspKeyCredentialThumbprint -KeyEntry @($application.keyCredentials) -Thumbprint $cert.Thumbprint)) {
            $steps.Add((New-MspStepResult -Step 'Certificate on existing application' -Status Warning -Detail 'The existing application does not have this certificate. Add it with Add-MspPartnerAppCertificate (it uses addKey so existing credentials are kept).'))
        }
    }

    $state.AppId = [string]$application.appId
    $state.ApplicationObjectId = [string]$application.id

    # 7. Read the application back
    $appPath = "applications/{0}?`$select=id,appId,displayName,signInAudience,web,requiredResourceAccess,keyCredentials" -f $application.id
    $readback = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition { Invoke-MspPartnerGraph @g -Path $appPath }
    if (-not $readback.Satisfied) {
        $steps.Add((New-MspStepResult -Step 'Verify application' -Status Failed -Detail "Could not read the application back. $($readback.LastError)"))
        return (& $finish)
    }
    $read = $readback.Value
    $issues = New-Object System.Collections.Generic.List[string]
    if ($read.signInAudience -ne 'AzureADMultipleOrgs') { $issues.Add("signInAudience is $($read.signInAudience)") }
    foreach ($uri in $RedirectUri) { if (@($read.web.redirectUris) -notcontains $uri) { $issues.Add("redirect URI $uri missing") } }
    $readPairs = @(foreach ($r in @($read.requiredResourceAccess)) { foreach ($a in @($r.resourceAccess)) { ('{0}|{1}' -f $r.resourceAppId, $a.id).ToLowerInvariant() } })
    foreach ($r in @($desiredRra)) { foreach ($a in @($r.resourceAccess)) { if ($readPairs -notcontains ('{0}|{1}' -f $r.resourceAppId, $a.id).ToLowerInvariant()) { $issues.Add("permission $($a.id) missing") } } }
    if ($cert -and $steps.Step -contains 'Create application' -and -not (Test-MspKeyCredentialThumbprint -KeyEntry @($read.keyCredentials) -Thumbprint $cert.Thumbprint)) { $issues.Add("certificate $($cert.Thumbprint) not found on the application") }
    if ($issues.Count -gt 0) { $steps.Add((New-MspStepResult -Step 'Verify application' -Status Failed -Detail ($issues -join '; '))) }
    else { $steps.Add((New-MspStepResult -Step 'Verify application' -Status Passed -Detail 'Audience, redirect URIs, permissions and certificate confirmed.')) }

    # 8. Service principal in the partner tenant
    $spPath = "servicePrincipals?`$filter=appId eq '{0}'&`$select=id,appId,appRoleAssignmentRequired" -f $state.AppId
    $sp = $null
    try { $sp = @(Invoke-MspPartnerGraph @g -Path $spPath) | Select-Object -First 1 } catch { $sp = $null }
    if ($sp) {
        $steps.Add((New-MspStepResult -Step 'Create service principal' -Status Passed -Detail "Service principal $($sp.id) already exists."))
    }
    elseif ($PSCmdlet.ShouldProcess("partner tenant $PartnerTenantId", "Create service principal for $($state.AppId)")) {
        try {
            $null = Invoke-MspPartnerGraph @g -Method POST -Path 'servicePrincipals' -Body @{ appId = $state.AppId }
            $wait = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition { @(Invoke-MspPartnerGraph @g -Path $spPath) | Select-Object -First 1 }
            if ($wait.Satisfied) {
                $sp = $wait.Value
                $steps.Add((New-MspStepResult -Step 'Create service principal' -Status Changed -Detail "Created and confirmed service principal $($sp.id)."))
            }
            else {
                $steps.Add((New-MspStepResult -Step 'Create service principal' -Status Failed -Detail "Created, but it could not be read back. $($wait.LastError)"))
            }
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Create service principal' -Status Failed -Detail $_.Exception.Message))
        }
    }
    else {
        $steps.Add((New-MspStepResult -Step 'Create service principal' -Status WhatIf -Detail 'Would create the service principal in the partner tenant.'))
    }
    if ($sp) { $state.ServicePrincipalId = [string]$sp.id }

    # 9. Admin consent in the partner tenant (delegated scopes only)
    if ($SkipAdminConsent) {
        $steps.Add((New-MspStepResult -Step 'Grant admin consent in partner tenant' -Status Skipped -Detail '-SkipAdminConsent was used.'))
    }
    elseif (-not $sp) {
        $steps.Add((New-MspStepResult -Step 'Grant admin consent in partner tenant' -Status Skipped -Detail 'No service principal yet.'))
    }
    else {
        foreach ($resource in $manifest.RequiredResourceAccess) {
            $stepName = "Grant admin consent: $($resource.resourceDisplayName)"
            $scopes = @($resource.resourceAccess | Where-Object { $_.type -eq 'Scope' } | ForEach-Object { $_.value })
            $roles = @($resource.resourceAccess | Where-Object { $_.type -eq 'Role' })
            if ($roles.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step "Application permissions: $($resource.resourceDisplayName)" -Status Skipped -Detail "$($roles.Count) application permissions are not granted in the partner tenant. Grant them per customer (Enable-MspExchangeAppAccess)."))
            }
            if ($scopes.Count -eq 0) { continue }
            $resourceSp = $resourceSps[$resource.resourceAppId]
            $grantPath = "oauth2PermissionGrants?`$filter=clientId eq '{0}' and consentType eq 'AllPrincipals' and resourceId eq '{1}'" -f $sp.id, $resourceSp.id
            try {
                $grant = @(Invoke-MspPartnerGraph @g -Path $grantPath) | Select-Object -First 1
                $current = if ($grant) { @(([string]$grant.scope).Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) } else { @() }
                $missing = @($scopes | Where-Object { $current -notcontains $_ })
                if ($missing.Count -eq 0) {
                    $steps.Add((New-MspStepResult -Step $stepName -Status Passed -Detail "All $($scopes.Count) scopes already consented."))
                    continue
                }
                if (-not $PSCmdlet.ShouldProcess("partner tenant $PartnerTenantId", "Grant tenant-wide consent for $($resource.resourceDisplayName): $($missing -join ' ')")) {
                    $steps.Add((New-MspStepResult -Step $stepName -Status WhatIf -Detail "Would consent: $($missing -join ' ')"))
                    continue
                }
                $union = @($current + $missing | Select-Object -Unique) -join ' '
                if ($grant) {
                    $null = Invoke-MspPartnerGraph @g -Method PATCH -Path "oauth2PermissionGrants/$($grant.id)" -Body @{ scope = $union }
                }
                else {
                    $null = Invoke-MspPartnerGraph @g -Method POST -Path 'oauth2PermissionGrants' -Body @{ clientId = $sp.id; consentType = 'AllPrincipals'; resourceId = $resourceSp.id; scope = $union }
                }
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                    $g2 = @(Invoke-MspPartnerGraph @g -Path $grantPath) | Select-Object -First 1
                    if ($g2) {
                        $have = @(([string]$g2.scope).Split(' ', [StringSplitOptions]::RemoveEmptyEntries))
                        if (@($scopes | Where-Object { $have -notcontains $_ }).Count -eq 0) { $g2 }
                    }
                }
                if ($check.Satisfied) { $steps.Add((New-MspStepResult -Step $stepName -Status Changed -Detail "Consented and confirmed: $($missing -join ' ')")) }
                else { $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail "Consent was sent but the readback does not show every scope. $($check.LastError)")) }
            }
            catch {
                $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail $_.Exception.Message))
            }
        }
    }

    # 10. Optional: only one group of technicians can sign in to the app in the partner tenant
    if ($RestrictToGroupId -and $sp) {
        $stepName = 'Restrict sign-in to technician group'
        if ($PSCmdlet.ShouldProcess("service principal $($sp.id)", "Require assignment and assign group $RestrictToGroupId")) {
            try {
                if (-not $sp.appRoleAssignmentRequired) {
                    $null = Invoke-MspPartnerGraph @g -Method PATCH -Path "servicePrincipals/$($sp.id)" -Body @{ appRoleAssignmentRequired = $true }
                }
                $assignPath = "servicePrincipals/$($sp.id)/appRoleAssignedTo"
                $assigned = @(Invoke-MspPartnerGraph @g -Path $assignPath | Where-Object { $_.principalId -eq $RestrictToGroupId })
                if ($assigned.Count -eq 0) {
                    $null = Invoke-MspPartnerGraph @g -Method POST -Path $assignPath -Body @{ principalId = $RestrictToGroupId; resourceId = $sp.id; appRoleId = '00000000-0000-0000-0000-000000000000' }
                }
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                    $spNow = @(Invoke-MspPartnerGraph @g -Path $spPath) | Select-Object -First 1
                    $groupNow = @(Invoke-MspPartnerGraph @g -Path $assignPath | Where-Object { $_.principalId -eq $RestrictToGroupId })
                    ($spNow.appRoleAssignmentRequired -eq $true) -and ($groupNow.Count -gt 0)
                }
                if ($check.Satisfied) { $steps.Add((New-MspStepResult -Step $stepName -Status Changed -Detail "Only members of $RestrictToGroupId can sign in to the app in the partner tenant.")) }
                else { $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail "Could not confirm the assignment requirement and group assignment. $($check.LastError)")) }
            }
            catch {
                $steps.Add((New-MspStepResult -Step $stepName -Status Failed -Detail $_.Exception.Message))
            }
        }
        else {
            $steps.Add((New-MspStepResult -Step $stepName -Status WhatIf -Detail "Would require assignment and assign group $RestrictToGroupId."))
        }
    }

    $thumbText = if ($state.CertificateThumbprint) { $state.CertificateThumbprint } else { '<Thumbprint>' }
    $state.NextSteps = @(
        ("Set-MspConfiguration -PartnerTenantId '{0}' -AppId '{1}' -CertificateThumbprint '{2}' -VaultName '<VaultName>'" -f $PartnerTenantId, $state.AppId, $thumbText),
        "Register-MspPartnerToken -UserPrincipalName '<technician admin UPN>'   (once per technician, MFA sign-in)",
        "Grant-MspPartnerAppConsent -TenantId '<customer tenant>' -WhatIf   (then without -WhatIf, then Test-MspPartnerAppConsent)",
        'Keep the certificate private key non-exportable. Never add credentials to the app''s service principal in customer tenants.'
    )
    foreach ($line in $state.NextSteps) { Write-Information -MessageData ("Next: {0}" -f $line) }
    & $finish
}
