#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Structural checks on the shipped partner app permission manifests. These read the JSON files only.
# They do not call Microsoft Graph, so they cannot prove an id resolves to its value. New-MspPartnerApp
# does that against the resource service principal before it writes anything.

BeforeDiscovery {
    $manifestRoot = Join-Path $PSScriptRoot '..' 'manifests'
    $script:PartnerManifests = foreach ($file in Get-ChildItem -Path $manifestRoot -Filter 'partner-app.*.json' -File) {
        @{ Name = $file.Name; Path = $file.FullName }
    }
}

BeforeAll {
    $script:ManifestRoot = Join-Path $PSScriptRoot '..' 'manifests'
    $script:GuidPattern = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'

    function Get-TestManifestEntry {
        param([Parameter(Mandatory)][string]$Path)
        $manifest = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
        foreach ($resource in $manifest.requiredResourceAccess) {
            foreach ($access in $resource.resourceAccess) {
                [pscustomobject]@{
                    ResourceAppId = $resource.resourceAppId
                    Id            = $access.id
                    Type          = $access.type
                    Value         = $access.value
                    Key           = '{0}|{1}|{2}' -f $resource.resourceAppId, $access.id, $access.type
                }
            }
        }
    }
}

Describe 'Partner app manifest <Name>' -ForEach $script:PartnerManifests {
    BeforeAll { $script:Entries = @(Get-TestManifestEntry -Path $Path) }

    It 'has at least one permission' {
        $script:Entries.Count | Should -BeGreaterThan 0
    }
    It 'uses a lower-case GUID for every resourceAppId and permission id' {
        foreach ($entry in $script:Entries) {
            $entry.ResourceAppId | Should -MatchExactly $script:GuidPattern -Because "$($entry.Value) has resourceAppId '$($entry.ResourceAppId)'"
            $entry.Id | Should -MatchExactly $script:GuidPattern -Because "$($entry.Value) has id '$($entry.Id)'"
        }
    }
    It 'holds delegated permissions only (type Scope)' {
        foreach ($entry in $script:Entries) {
            $entry.Type | Should -BeExactly 'Scope' -Because "$($entry.Value) ($($entry.Id)) must be delegated"
        }
    }
    It 'documents every permission with a value' {
        foreach ($entry in $script:Entries) {
            $entry.Value | Should -Not -BeNullOrEmpty -Because "id $($entry.Id) needs a value"
        }
    }
}

Describe 'partner-app.full.json' {
    BeforeAll {
        $script:Minimal = @(Get-TestManifestEntry -Path (Join-Path $script:ManifestRoot 'partner-app.minimal.json'))
        $script:Full = @(Get-TestManifestEntry -Path (Join-Path $script:ManifestRoot 'partner-app.full.json'))
    }

    It 'is a strict superset of partner-app.minimal.json (same resourceAppId, id and type)' {
        $fullKeys = [System.Collections.Generic.HashSet[string]]::new([string[]]@($script:Full.Key), [System.StringComparer]::OrdinalIgnoreCase)
        $missing = @($script:Minimal | Where-Object { -not $fullKeys.Contains($_.Key) } | ForEach-Object { $_.Value })
        $missing | Should -BeNullOrEmpty -Because 'moving from the minimal to the full manifest must never remove a permission'
        $script:Full.Count | Should -BeGreaterThan $script:Minimal.Count
    }
    It 'gives each shared permission the same value as the minimal manifest' {
        $fullByKey = @{}
        foreach ($entry in $script:Full) { $fullByKey[$entry.Key] = $entry.Value }
        foreach ($entry in $script:Minimal) {
            $fullByKey[$entry.Key] | Should -BeExactly $entry.Value
        }
    }
}
