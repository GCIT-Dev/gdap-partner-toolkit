function Resolve-MspResource {
    <#
    .SYNOPSIS
        Normalises a resource name, URI or application ID to a canonical form.
    .DESCRIPTION
        Accepts a friendly alias (Graph, Exchange, PartnerCenter, ManagementApi,
        Defender, AzureManagement, TeamsAdmin), an https or api:// resource URI
        (trailing slash and /.default are removed) or a resource application
        ID GUID. Returns the canonical lower-case URI used in cache keys, the
        resource app ID when known, and every audience form accepted when a
        token's aud claim is compared.
    .PARAMETER Resource
        The resource to resolve.
    .EXAMPLE
        Resolve-MspResource -Resource Graph
    .EXAMPLE
        Resolve-MspResource -Resource 'https://contoso.sharepoint.com/'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Resource
    )

    # Public first-party resource URIs and application IDs.
    $known = @(
        @{ Name = 'Graph'; Uri = 'https://graph.microsoft.com'; AppId = '00000003-0000-0000-c000-000000000000'; Extra = @('https://graph.microsoft.com/') }
        @{ Name = 'Exchange'; Uri = 'https://outlook.office365.com'; AppId = '00000002-0000-0ff1-ce00-000000000000'; Extra = @('https://outlook.office.com') }
        @{ Name = 'PartnerCenter'; Uri = 'https://api.partnercenter.microsoft.com'; AppId = 'fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd'; Extra = @() }
        @{ Name = 'ManagementApi'; Uri = 'https://manage.office.com'; AppId = 'c5393580-f805-4401-95e8-94b7a6ef2fc2'; Extra = @() }
        @{ Name = 'Defender'; Uri = 'https://api.securitycenter.microsoft.com'; AppId = 'fc780465-2017-40d4-a0c5-307022471b92'; Extra = @('https://securitycenter.onmicrosoft.com/windowsatpservice') }
        @{ Name = 'AzureManagement'; Uri = 'https://management.azure.com'; AppId = '797f4846-ba00-4fd7-ba43-dac1f8f63013'; Extra = @('https://management.core.windows.net') }
        @{ Name = 'TeamsAdmin'; Uri = '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'; AppId = '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'; Extra = @() }
    )
    $aliases = @{
        'graph'                 = 'Graph'
        'msgraph'               = 'Graph'
        'microsoftgraph'        = 'Graph'
        'exchange'              = 'Exchange'
        'exchangeonline'        = 'Exchange'
        'partnercenter'         = 'PartnerCenter'
        'managementapi'         = 'ManagementApi'
        'office365managementapi' = 'ManagementApi'
        'defender'              = 'Defender'
        'defenderforendpoint'   = 'Defender'
        'azuremanagement'       = 'AzureManagement'
        'arm'                   = 'AzureManagement'
        'teamsadmin'            = 'TeamsAdmin'
        'teams'                 = 'TeamsAdmin'
    }

    $value = $Resource.Trim()
    if ($value -match '/\.default$') { $value = $value.Substring(0, $value.Length - '/.default'.Length) }
    $value = $value.TrimEnd('/')
    $lower = $value.ToLowerInvariant()

    $entry = $null
    if ($aliases.ContainsKey($lower)) {
        $entry = $known | Where-Object { $_.Name -eq $aliases[$lower] }
    }
    else {
        $entry = $known | Where-Object { $_.Uri -eq $lower -or $_.AppId -eq $lower -or $_.Extra -contains $lower } | Select-Object -First 1
    }

    if ($entry) {
        $audiences = @($entry.Uri, "$($entry.Uri)/", $entry.AppId) + @($entry.Extra)
        return [pscustomobject]@{
            Name      = $entry.Name
            Uri       = $entry.Uri
            AppId     = $entry.AppId
            Audiences = @($audiences | Select-Object -Unique)
        }
    }

    if ($lower -match '^https://[a-z0-9-]+(-admin)?\.sharepoint\.com$') {
        return [pscustomobject]@{
            Name      = 'SharePoint'
            Uri       = $lower
            AppId     = '00000003-0000-0ff1-ce00-000000000000'
            Audiences = @($lower, "$lower/", '00000003-0000-0ff1-ce00-000000000000')
        }
    }
    if ($lower -match $script:MspConstants.GuidPattern) {
        return [pscustomobject]@{ Name = 'Custom'; Uri = $lower; AppId = $lower; Audiences = @($lower) }
    }
    if ($lower -match '^(https|api)://[^\s/?#]+(/[^\s?#]*)?$') {
        return [pscustomobject]@{ Name = 'Custom'; Uri = $lower; AppId = $null; Audiences = @($lower, "$lower/") }
    }

    throw (New-MspErrorRecord -Message "Resource '$Resource' is not a known alias, an https:// or api:// URI, or an application ID." -ErrorId 'MspGdap.Resource.Invalid' -Category InvalidArgument -TargetObject $Resource)
}
