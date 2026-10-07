function Get-MspClientCredentialBody {
    <#
    .SYNOPSIS
        Returns the client authentication fields for a token request.
    .DESCRIPTION
        For a certificate credential this is client_assertion_type plus a fresh
        client_assertion whose aud is the token endpoint of -TenantId. For a
        client secret (discouraged) it is client_secret read from the vault.
    .PARAMETER Configuration
        The active configuration object.
    .PARAMETER TenantId
        Tenant whose token endpoint will receive the request.
    .EXAMPLE
        $credential = Get-MspClientCredentialBody -Configuration $config -TenantId $customerTenantId
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Configuration,

        [Parameter(Mandatory)]
        [string]$TenantId
    )

    if ($Configuration.CredentialType -eq 'ClientSecret') {
        $name = Get-MspClientSecretName -AppId $Configuration.AppId
        try {
            $secret = Get-Secret -Name $name -Vault $Configuration.VaultName -ErrorAction Stop
        }
        catch {
            throw (New-MspErrorRecord -Message "The partner app client secret ($name) was not found in vault '$($Configuration.VaultName)'. Run Set-MspConfiguration -ClientSecret, or better, switch to a certificate." -ErrorId 'MspGdap.ClientSecret.NotFound' -Category ObjectNotFound -TargetObject $name)
        }
        $plain = if ($secret -is [securestring]) { ConvertFrom-MspSecureString -SecureString $secret } else { [string]$secret }
        return [ordered]@{ client_secret = $plain }
    }

    $certificate = Get-MspSigningCertificate -Configuration $Configuration
    $algorithm = if ($Configuration.SigningAlgorithm) { $Configuration.SigningAlgorithm } else { 'PS256' }
    $assertion = New-MspClientAssertion -Certificate $certificate -ClientId $Configuration.AppId -TenantId $TenantId -Algorithm $algorithm
    [ordered]@{
        client_assertion_type = $script:MspConstants.ClientAssertionType
        client_assertion      = $assertion
    }
}
