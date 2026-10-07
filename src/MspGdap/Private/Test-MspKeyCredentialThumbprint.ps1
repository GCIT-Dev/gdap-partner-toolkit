function Test-MspKeyCredentialThumbprint {
    <#
    .SYNOPSIS
        True when one of the application's keyCredentials is the certificate with this SHA-1 thumbprint.
    .DESCRIPTION
        Microsoft Graph v1.0 returns customKeyIdentifier for an uploaded certificate as the hex thumbprint
        (for example 2CEF3E3B...). Some older credentials and some clients store it as the Base64 of the
        thumbprint bytes instead. A 40 character hex thumbprint is also valid Base64, so the hex form must be
        compared first: decoding it as Base64 succeeds and yields different bytes, which would hide a match.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$KeyEntry,
        [Parameter(Mandatory)][string]$Thumbprint
    )
    foreach ($credential in @($KeyEntry)) {
        if ($null -eq $credential) { continue }
        $identifier = [string]$credential.customKeyIdentifier
        if (-not $identifier) { continue }
        if ($identifier -ieq $Thumbprint) { return $true }
        try {
            $hex = ([BitConverter]::ToString([Convert]::FromBase64String($identifier))).Replace('-', '')
            if ($hex -ieq $Thumbprint) { return $true }
        }
        catch {
            Write-Verbose 'customKeyIdentifier is neither the hex thumbprint nor Base64 of it.'
        }
    }
    return $false
}
