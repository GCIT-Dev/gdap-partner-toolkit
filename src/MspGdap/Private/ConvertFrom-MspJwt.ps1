function ConvertFrom-MspJwt {
    <#
    .SYNOPSIS
        Decodes the header and claims of a JWT. Does NOT validate the signature.
    .DESCRIPTION
        This is a convenience decoder for diagnostics and best-effort checks
        (exp, aud, tid, scp, roles, amr). It performs no signature validation
        and must never be used to make a trust decision about a token received
        from somewhere else.

        Microsoft states that clients should treat access tokens as opaque, and
        some access tokens (for example some Microsoft Graph tokens) are not
        decodable JWTs. This function therefore never throws on a token it
        cannot decode. It returns $null instead.
    .PARAMETER Token
        The token as a string or SecureString.
    .OUTPUTS
        PSCustomObject with Header and Claims properties, or $null.
    .EXAMPLE
        (ConvertFrom-MspJwt -Token $idToken).Claims.preferred_username
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Token
    )
    process {
        if ($null -eq $Token) { return $null }
        $text = if ($Token -is [securestring]) { ConvertFrom-MspSecureString -SecureString $Token } else { [string]$Token }
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }

        $parts = $text.Split('.')
        if ($parts.Count -lt 2) {
            Write-Verbose 'Token is not a decodable JWT (opaque token). Claim checks skipped.'
            return $null
        }
        try {
            $headerJson = [System.Text.Encoding]::UTF8.GetString((ConvertFrom-MspBase64Url -Value $parts[0]))
            $claimsJson = [System.Text.Encoding]::UTF8.GetString((ConvertFrom-MspBase64Url -Value $parts[1]))
            $header = $headerJson | ConvertFrom-Json -ErrorAction Stop
            $claims = $claimsJson | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Verbose 'Token is not a decodable JWT (encrypted or opaque). Claim checks skipped.'
            return $null
        }
        if ($null -eq $claims -or $claims -isnot [pscustomobject]) { return $null }
        [pscustomobject]@{
            Header = $header
            Claims = $claims
        }
    }
}
