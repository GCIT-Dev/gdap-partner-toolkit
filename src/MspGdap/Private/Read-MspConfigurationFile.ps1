function Read-MspConfigurationFile {
    <#
    .SYNOPSIS
        Reads the non-secret MspGdap configuration JSON, or returns defaults.
    .PARAMETER Path
        Path to config.json.
    .EXAMPLE
        Read-MspConfigurationFile -Path "$HOME/.mspgdap/config.json"
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $config = [ordered]@{
        ConfigVersion            = 1
        PartnerTenantId          = $null
        AppId                    = $null
        CredentialType           = 'Certificate'
        CertificateThumbprint    = $null
        CertificateStoreLocation = 'CurrentUser'
        CertificatePath          = $null
        SigningAlgorithm         = 'PS256'
        VaultName                = $null
        LoopbackPort             = 0
        TechnicianUpn            = $null
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $stored = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw (New-MspErrorRecord -Message "The MspGdap configuration file '$Path' could not be read: $($_.Exception.Message)" -ErrorId 'MspGdap.Configuration.Unreadable' -Category ReadError -TargetObject $Path)
        }
        foreach ($property in $stored.PSObject.Properties) {
            if ($config.Contains($property.Name)) {
                $config[$property.Name] = $property.Value
            }
        }
    }
    [pscustomobject]$config
}
