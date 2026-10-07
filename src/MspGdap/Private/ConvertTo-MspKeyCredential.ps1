function ConvertTo-MspKeyCredential {
    <#
    .SYNOPSIS
        Builds a Microsoft Graph keyCredential (public key only) from a certificate.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)

    $name = $Certificate.Subject
    if ($name.Length -gt 90) { $name = $name.Substring(0, 90) }
    @{
        type        = 'AsymmetricX509Cert'
        usage       = 'Verify'
        key         = [Convert]::ToBase64String($Certificate.RawData)
        displayName = $name
    }
}
