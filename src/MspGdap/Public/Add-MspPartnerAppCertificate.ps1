function Add-MspPartnerAppCertificate {
    <#
    .SYNOPSIS
        Adds (or rotates in) a certificate credential on the partner app or automation app in the partner tenant.
    .DESCRIPTION
        Only the public key is uploaded. Two documented Graph paths:
          - The app has no valid certificate yet: Update application (PATCH keyCredentials).
          - The app already has a valid certificate: addKey, which keeps existing credentials and needs a
            proof-of-possession token signed by one of the existing certificates (-ProofCertificateThumbprint
            or -ProofCertificate, with its private key on this machine).
        PATCH replaces every existing key, so it is only used when no valid key exists, or with -ReplaceExisting.
        The new key is read back by thumbprint before the step reports Changed.
        Never adds credentials to a service principal (customer tenants included).
    .PARAMETER AppId
        Application (client) ID of the app in your partner tenant.
    .PARAMETER PartnerTenantId
        Tenant ID (GUID) of the partner tenant. The command stops if the Graph session is in any other tenant.
    .PARAMETER CertificateThumbprint
        Thumbprint of the new certificate in CurrentUser\My or LocalMachine\My. Only the public key is uploaded.
    .PARAMETER Certificate
        The new certificate as an X509Certificate2 object. Only the public key is uploaded.
    .PARAMETER CertificatePath
        Path to the new public certificate file (.cer, .crt or .pem). PFX files are refused.
    .PARAMETER ProofCertificateThumbprint
        Thumbprint of a certificate that is already valid on the app, with its private key on this machine.
        It signs the proof-of-possession token that addKey needs.
    .PARAMETER ProofCertificate
        The same as -ProofCertificateThumbprint, as an X509Certificate2 object with its private key.
    .PARAMETER ReplaceExisting
        Replace all existing certificate credentials with the new one (PATCH). The app stops working with the
        old certificate immediately. Use only when you have to.
    .PARAMETER Transport
        MgGraph (default, an existing Connect-MgGraph session in the partner tenant) or MspGdap (the
        technician's MspGdap partner tenant token).
    .PARAMETER ReadbackTimeoutSeconds
        How long to keep reading the application back before reporting the change as Failed. Default 60.
    .EXAMPLE
        Add-MspPartnerAppCertificate -AppId '<PartnerAppId>' -PartnerTenantId '<PartnerTenantId>' -CertificateThumbprint '<NewThumbprint>' -ProofCertificateThumbprint '<CurrentThumbprint>'
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Thumbprint')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$PartnerTenantId,
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$CertificateThumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory, ParameterSetName = 'Path')][string]$CertificatePath,
        [ValidatePattern('^[0-9a-fA-F]{40}$')][string]$ProofCertificateThumbprint,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$ProofCertificate,
        [switch]$ReplaceExisting,
        [ValidateSet('MgGraph', 'MspGdap')][string]$Transport = 'MgGraph',
        [ValidateRange(0, 600)][int]$ReadbackTimeoutSeconds = 60
    )
    $PartnerTenantId = $PartnerTenantId.ToLowerInvariant()
    $steps = New-Object System.Collections.Generic.List[object]
    $state = [ordered]@{ AppId = $AppId; ApplicationObjectId = $null; CertificateThumbprint = $null; Method = $null }
    $finish = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Add-MspPartnerAppCertificate' -TenantId $PartnerTenantId -Target $AppId -Steps $steps.ToArray() -Property $state }
    $g = @{ Transport = $Transport }

    $contextStep = Test-MspSetupContext -Transport $Transport -PartnerTenantId $PartnerTenantId -RequiredScope @('Application.ReadWrite.All')
    $steps.Add($contextStep)
    if ($contextStep.Status -eq 'Failed') { return (& $finish) }

    try {
        $cert = switch ($PSCmdlet.ParameterSetName) {
            'Thumbprint' { Resolve-MspCertificate -Thumbprint $CertificateThumbprint }
            'Certificate' { Resolve-MspCertificate -Certificate $Certificate }
            'Path' { Resolve-MspCertificate -Path $CertificatePath }
        }
        $state.CertificateThumbprint = $cert.Thumbprint
        $steps.Add((New-MspStepResult -Step 'Check certificate' -Status Passed -Detail ("{0}, expires {1:yyyy-MM-dd}." -f $cert.Thumbprint, $cert.NotAfter)))
        if ($cert.HasPrivateKey) {
            $privateKey = $null
            try { $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert) } catch { $privateKey = $null }
            if ($privateKey -is [System.Security.Cryptography.RSACng]) {
                $steps.Add((New-MspStepResult -Step 'Certificate key provider' -Status Warning -Detail 'CNG key. Fine for MspGdap sign-in, but Exchange app-only (Connect-ExchangeOnline -CertificateThumbprint) does not support CNG certificates. Use a CSP key for an automation app.'))
            }
            elseif ($privateKey -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
                $steps.Add((New-MspStepResult -Step 'Certificate key provider' -Status Warning -Detail 'Legacy CSP key. Works for Exchange app-only. Some CSP providers cannot sign PS256 client assertions, in which case MspGdap needs RS256 for this certificate.'))
            }
        }
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Check certificate' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }

    try {
        $app = @(Invoke-MspPartnerGraph @g -Path ("applications?`$filter=appId eq '{0}'&`$select=id,appId,displayName,keyCredentials" -f $AppId)) | Select-Object -First 1
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Find application' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }
    if (-not $app) {
        $steps.Add((New-MspStepResult -Step 'Find application' -Status Failed -Detail "Application $AppId was not found in partner tenant $PartnerTenantId."))
        return (& $finish)
    }
    $state.ApplicationObjectId = [string]$app.id
    $steps.Add((New-MspStepResult -Step 'Find application' -Status Passed -Detail "$($app.displayName) has $(@($app.keyCredentials).Count) key credentials."))

    if (Test-MspKeyCredentialThumbprint -KeyEntry @($app.keyCredentials) -Thumbprint $cert.Thumbprint) {
        $steps.Add((New-MspStepResult -Step 'Add certificate' -Status Passed -Detail 'This certificate is already on the application. Nothing to do.'))
        return (& $finish)
    }

    $now = [datetime]::UtcNow
    $validKeys = @($app.keyCredentials | Where-Object { $_.type -eq 'AsymmetricX509Cert' -and $_.endDateTime -and ([datetime]$_.endDateTime).ToUniversalTime() -gt $now })
    $keyCredential = ConvertTo-MspKeyCredential -Certificate $cert
    $appPath = "applications/{0}?`$select=id,keyCredentials" -f $app.id

    if ($validKeys.Count -eq 0 -or $ReplaceExisting) {
        $state.Method = 'PATCH keyCredentials'
        $action = if ($validKeys.Count -eq 0) { 'Add the first valid certificate (PATCH, removes expired keys)' } else { "REPLACE all $($validKeys.Count) valid certificates (PATCH)" }
        if (-not $PSCmdlet.ShouldProcess("application $AppId", $action)) {
            $steps.Add((New-MspStepResult -Step 'Add certificate' -Status WhatIf -Detail $action))
            return (& $finish)
        }
        try { $null = Invoke-MspPartnerGraph @g -Method PATCH -Path "applications/$($app.id)" -Body @{ keyCredentials = @($keyCredential) } }
        catch {
            $steps.Add((New-MspStepResult -Step 'Add certificate' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }
    }
    else {
        $state.Method = 'addKey'
        $proof = $null
        try {
            if ($ProofCertificate) { $proofCert = Resolve-MspCertificate -Certificate $ProofCertificate -RequirePrivateKey }
            elseif ($ProofCertificateThumbprint) { $proofCert = Resolve-MspCertificate -Thumbprint $ProofCertificateThumbprint -RequirePrivateKey }
            else { throw "The application already has $($validKeys.Count) valid certificates. Pass -ProofCertificateThumbprint (an existing certificate with its private key) so the new one is added with addKey, or use -ReplaceExisting." }
            if (-not (Test-MspKeyCredentialThumbprint -KeyEntry $validKeys -Thumbprint $proofCert.Thumbprint)) {
                throw "Proof certificate $($proofCert.Thumbprint) is not one of the application's valid certificates."
            }
            $proof = New-MspAddKeyProof -Certificate $proofCert -ApplicationObjectId $app.id
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Build proof of possession' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }
        $steps.Add((New-MspStepResult -Step 'Build proof of possession' -Status Passed -Detail "Signed with existing certificate $($proofCert.Thumbprint)."))
        if (-not $PSCmdlet.ShouldProcess("application $AppId", "Add certificate $($cert.Thumbprint) with addKey (existing certificates are kept)")) {
            $steps.Add((New-MspStepResult -Step 'Add certificate' -Status WhatIf -Detail 'Would call addKey.'))
            return (& $finish)
        }
        try {
            $null = Invoke-MspPartnerGraph @g -Method POST -Path "applications/$($app.id)/addKey" -Body @{ keyCredential = @{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = $keyCredential.key }; passwordCredential = $null; proof = $proof }
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Add certificate' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }
        finally { $proof = $null }
    }

    $newThumbprint = $cert.Thumbprint
    $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
        $now2 = Invoke-MspPartnerGraph @g -Path $appPath
        Test-MspKeyCredentialThumbprint -KeyEntry @($now2.keyCredentials) -Thumbprint $newThumbprint
    }
    if ($check.Satisfied) { $steps.Add((New-MspStepResult -Step 'Add certificate' -Status Changed -Detail "Certificate $newThumbprint added with $($state.Method) and confirmed by readback.")) }
    else { $steps.Add((New-MspStepResult -Step 'Add certificate' -Status Failed -Detail "The request was accepted but certificate $newThumbprint was not found on readback. $($check.LastError)")) }
    & $finish
}
