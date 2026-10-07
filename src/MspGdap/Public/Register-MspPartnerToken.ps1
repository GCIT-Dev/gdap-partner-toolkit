function Register-MspPartnerToken {
    <#
    .SYNOPSIS
        Signs a technician in once (authorisation code with PKCE and MFA) and
        stores their delegated refresh token in the SecretManagement vault.
    .DESCRIPTION
        This is the Secure Application Model consent step for one technician.

        1. Opens the browser at the partner tenant's v2.0 /authorize endpoint
           for the partner app, with PKCE (S256), a random state and nonce, and
           scopes "<resource>/.default offline_access openid profile".
        2. Receives the code on a one-shot listener at http://localhost:<port>
           (register http://localhost as a Web redirect URI on the partner app).
           Device code flow is deliberately not used because Microsoft
           recommends blocking it with Conditional Access.
        3. Redeems the code with the partner app's certificate (or client
           secret) and the PKCE verifier.
        4. Checks the sign-in: the account belongs to the partner tenant, it
           matches -UserPrincipalName when given, an id_token was returned and
           its nonce matches this sign-in, and MFA was performed (amr contains
           mfa). Partner Center rejects App+User calls without an MFA claim, so
           a token is not stored when amr shows no MFA, or when neither token
           exposes amr at all, unless -SkipMfaCheck is used.
        5. Stores the refresh token in the vault as
           MspGdap-<appId>-<hash of UPN> with non-secret metadata. The UPN is not
           in the secret name.

        Use a dedicated partner admin account that is a member of your GDAP
        security groups (and AdminAgents for Partner Center work). The output
        never includes token values.
    .PARAMETER UserPrincipalName
        Expected technician account. Used as login_hint and checked after sign-in.
    .PARAMETER Port
        Loopback port for the redirect. Overrides the configured LoopbackPort.
        0 picks a free port.
    .PARAMETER TimeoutSeconds
        How long to wait for the browser sign-in. Default 300.
    .PARAMETER NoBrowser
        Print the sign-in URL instead of opening a browser.
    .PARAMETER SkipMfaCheck
        Store the token even when the sign-in did not report MFA, or MFA could not
        be confirmed from the token claims. Not recommended.
    .PARAMETER SetAsDefault
        Save this technician as TechnicianUpn in the configuration file.
    .OUTPUTS
        MspGdap.TechnicianRegistration
    .EXAMPLE
        Register-MspPartnerToken -UserPrincipalName 'tech-admin@contoso.onmicrosoft.com' -SetAsDefault
    .EXAMPLE
        Register-MspPartnerToken -NoBrowser -Port 53682
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    [OutputType('MspGdap.TechnicianRegistration')]
    param(
        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
        [string]$UserPrincipalName,

        [ValidateRange(0, 65535)]
        [int]$Port,

        [ValidateRange(30, 1800)]
        [int]$TimeoutSeconds = 300,

        [switch]$NoBrowser,

        [switch]$SkipMfaCheck,

        [switch]$SetAsDefault
    )

    try {
        $config = Get-MspConfigurationInternal -RequireComplete
        Assert-MspSecretManagement -VaultName $config.VaultName
        $partnerTenant = ([string]$config.PartnerTenantId).ToLowerInvariant()
        $appId = ([string]$config.AppId).ToLowerInvariant()

        $target = if ($UserPrincipalName) { $UserPrincipalName } else { 'the technician who signs in' }
        if (-not $PSCmdlet.ShouldProcess("vault '$($config.VaultName)'", "Sign in $target to partner app $appId and store the refresh token")) {
            return
        }

        $listenPort = if ($PSBoundParameters.ContainsKey('Port')) { $Port } elseif ($config.LoopbackPort) { [int]$config.LoopbackPort } else { 0 }
        if ($listenPort -le 0) { $listenPort = Get-MspFreeLoopbackPort }
        $redirectUri = "http://localhost:$listenPort"

        $pkce = New-MspPkceChallenge
        $state = New-MspRandomString
        $nonce = New-MspRandomString
        $scope = 'https://graph.microsoft.com/.default offline_access openid profile'

        $query = [ordered]@{
            client_id             = $appId
            response_type         = 'code'
            redirect_uri          = $redirectUri
            response_mode         = 'query'
            scope                 = $scope
            state                 = $state
            nonce                 = $nonce
            code_challenge        = $pkce.Challenge
            code_challenge_method = 'S256'
            prompt                = 'select_account'
        }
        if ($UserPrincipalName) { $query.login_hint = $UserPrincipalName }
        $queryString = ($query.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, [System.Uri]::EscapeDataString([string]$_.Value) }) -join '&'
        $authorizationUrl = '{0}/{1}/oauth2/v2.0/authorize?{2}' -f $script:MspConstants.Authority, $partnerTenant, $queryString

        $callback = Start-MspLoopbackListener -Port $listenPort -ExpectedState $state -AuthorizationUrl $authorizationUrl -TimeoutSeconds $TimeoutSeconds -NoBrowser:$NoBrowser

        $body = [ordered]@{
            client_id     = $appId
            grant_type    = 'authorization_code'
            code          = $callback.Code
            redirect_uri  = $callback.RedirectUri
            code_verifier = $pkce.Verifier
            scope         = $scope
        }
        $credential = Get-MspClientCredentialBody -Configuration $config -TenantId $partnerTenant
        foreach ($key in $credential.Keys) { $body[$key] = $credential[$key] }

        $requestedAt = [DateTimeOffset]::UtcNow
        $response = Invoke-MspTokenRequest -TenantId $partnerTenant -Body $body
        $body = $null
        $callback = $null

        if (-not $response.PSObject.Properties['refresh_token'] -or -not $response.refresh_token) {
            throw (New-MspErrorRecord -Message 'The sign-in did not return a refresh token. Check that offline_access is consented for the partner app in the partner tenant.' -ErrorId 'MspGdap.Register.NoRefreshToken' -Category InvalidResult -TargetObject $appId)
        }

        $idClaims = $null
        if ($response.PSObject.Properties['id_token'] -and $response.id_token) {
            $decodedId = ConvertFrom-MspJwt -Token $response.id_token
            if ($decodedId) { $idClaims = $decodedId.Claims }
        }
        $decodedAccess = ConvertFrom-MspJwt -Token $response.access_token
        $accessClaims = if ($decodedAccess) { $decodedAccess.Claims } else { $null }

        # openid is always requested, so an id_token carrying this sign-in's nonce must come back.
        if (-not $idClaims) {
            throw (New-MspErrorRecord -Message 'The sign-in did not return a readable id_token, so it cannot be tied to this request. Nothing was stored.' -ErrorId 'MspGdap.Register.NoIdToken' -Category SecurityError -TargetObject $appId)
        }
        if (-not $idClaims.PSObject.Properties['nonce'] -or [string]$idClaims.nonce -cne $nonce) {
            throw (New-MspErrorRecord -Message 'The id_token nonce is missing or does not match this sign-in. Nothing was stored.' -ErrorId 'MspGdap.Register.NonceMismatch' -Category SecurityError -TargetObject $appId)
        }

        $signedInUpn = $null
        foreach ($claimSet in @($idClaims, $accessClaims)) {
            if (-not $claimSet) { continue }
            foreach ($name in 'preferred_username', 'upn', 'unique_name') {
                if (-not $signedInUpn -and $claimSet.PSObject.Properties[$name] -and $claimSet.$name) { $signedInUpn = [string]$claimSet.$name }
            }
        }
        if (-not $signedInUpn) { $signedInUpn = $UserPrincipalName }
        if (-not $signedInUpn) {
            throw (New-MspErrorRecord -Message 'The signed-in account could not be identified from the token response. Run again with -UserPrincipalName.' -ErrorId 'MspGdap.Register.UnknownUser' -Category InvalidResult -TargetObject $appId)
        }
        $signedInUpn = $signedInUpn.Trim().ToLowerInvariant()
        if ($UserPrincipalName -and $signedInUpn -ne $UserPrincipalName.Trim().ToLowerInvariant()) {
            throw (New-MspErrorRecord -Message "You signed in as $signedInUpn but -UserPrincipalName was $UserPrincipalName. Nothing was stored." -ErrorId 'MspGdap.Register.UserMismatch' -Category SecurityError -TargetObject $signedInUpn)
        }

        $tokenTenant = $null
        foreach ($claimSet in @($idClaims, $accessClaims)) {
            if (-not $tokenTenant -and $claimSet -and $claimSet.PSObject.Properties['tid'] -and $claimSet.tid) { $tokenTenant = ([string]$claimSet.tid).ToLowerInvariant() }
        }
        if ($tokenTenant -and $tokenTenant -ne $partnerTenant) {
            throw (New-MspErrorRecord -Message "The sign-in was for tenant $tokenTenant, not the partner tenant $partnerTenant. Use a partner tenant admin account. Nothing was stored." -ErrorId 'MspGdap.Register.WrongTenant' -Category SecurityError -TargetObject $tokenTenant)
        }

        $amr = @()
        foreach ($claimSet in @($accessClaims, $idClaims)) {
            if ($amr.Count -eq 0 -and $claimSet -and $claimSet.PSObject.Properties['amr'] -and $claimSet.amr) { $amr = @($claimSet.amr) }
        }
        $mfaConfirmed = $null
        if ($amr.Count -gt 0) {
            $mfaConfirmed = $amr -contains 'mfa'
            if (-not $mfaConfirmed -and -not $SkipMfaCheck) {
                throw (New-MspErrorRecord -Message "The sign-in for $signedInUpn did not include multifactor authentication (amr: $($amr -join ', ')). Partner Center and GDAP delegated access need an MFA-backed refresh token. Enforce MFA for this account and try again. Nothing was stored." -ErrorId 'MspGdap.Register.MfaMissing' -Category SecurityError -TargetObject $signedInUpn)
            }
            if (-not $mfaConfirmed) {
                Write-Warning "Storing a refresh token for $signedInUpn without MFA because -SkipMfaCheck was used. Partner Center calls will fail with 'MFA required'."
            }
        }
        elseif (-not $SkipMfaCheck) {
            throw (New-MspErrorRecord -Message "MFA could not be confirmed for $signedInUpn because neither token exposed the amr claim. Partner Center and GDAP delegated access need an MFA-backed refresh token. Nothing was stored. If you are sure MFA is enforced for this account, run again with -SkipMfaCheck." -ErrorId 'MspGdap.Register.MfaUnverified' -Category SecurityError -TargetObject $signedInUpn)
        }
        else {
            Write-Warning "Storing a refresh token for $signedInUpn without MFA evidence because -SkipMfaCheck was used. Make sure MFA is enforced for this account in the partner tenant."
        }

        $secretName = Save-MspRefreshToken -UserPrincipalName $signedInUpn -AppId $appId -VaultName $config.VaultName -RefreshToken (ConvertTo-MspSecureString -Value $response.refresh_token)

        # Cache the partner-tenant Graph token from this sign-in.
        if ($response.PSObject.Properties['access_token'] -and $response.access_token) {
            $expiresIn = if ($response.PSObject.Properties['expires_in'] -and $response.expires_in) { [int]$response.expires_in } else { 0 }
            $scopes = @()
            if ($response.PSObject.Properties['scope'] -and $response.scope) {
                $scopes = @(([string]$response.scope).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { Get-MspScopeShortName -Scope $_ })
            }
            $key = Get-MspTokenCacheKey -TenantId $partnerTenant -Resource 'Graph' -AppId $appId
            $script:MspTokenCache[$key] = [pscustomobject]@{
                PSTypeName        = 'MspGdap.TokenCacheEntry'
                TenantId          = $partnerTenant
                Resource          = 'https://graph.microsoft.com'
                AppId             = $appId
                UserPrincipalName = $signedInUpn
                AccessToken       = ConvertTo-MspSecureString -Value $response.access_token
                ExpiresOn         = $requestedAt.AddSeconds($expiresIn)
                AcquiredOn        = $requestedAt
                Scopes            = $scopes
                Audience          = if ($accessClaims -and $accessClaims.PSObject.Properties['aud']) { [string]$accessClaims.aud } else { $null }
            }
        }
        $response = $null

        $script:MspState.CurrentUpn = $signedInUpn
        if ($SetAsDefault -or -not $config.TechnicianUpn) {
            Set-MspConfiguration -TechnicianUpn $signedInUpn -Confirm:$false -WhatIf:$false
        }

        [pscustomobject]@{
            PSTypeName        = 'MspGdap.TechnicianRegistration'
            UserPrincipalName = $signedInUpn
            AppId             = $appId
            PartnerTenantId   = $partnerTenant
            VaultName         = $config.VaultName
            SecretName        = $secretName
            MfaConfirmed      = $mfaConfirmed
            RegisteredOn      = $requestedAt
            RefreshTokenStored = $true
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
