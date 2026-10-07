#Requires -Version 7.4
# Offline smoke tests for scripts/automation. MspGdap calls, SecretManagement, IT Glue, UniFi and
# Power Automate calls are all mocked. Nothing here contacts Microsoft, a customer tenant or any
# third-party API.

BeforeAll {
    $script:RepoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..')).Path
    $script:ScriptRoot = Join-Path -Path $script:RepoRoot -ChildPath 'scripts/automation'
    $srcPath = Join-Path -Path $script:RepoRoot -ChildPath 'src'
    if (($env:PSModulePath -split [System.IO.Path]::PathSeparator) -notcontains $srcPath) {
        $env:PSModulePath = $srcPath + [System.IO.Path]::PathSeparator + $env:PSModulePath
    }
    Import-Module (Join-Path -Path $srcPath -ChildPath 'MspGdap/MspGdap.psd1') -Force

    $script:CustomerId = '22222222-2222-4222-8222-222222222222'
    $script:Tenant = 'contoso.onmicrosoft.com'
    $script:OwnScripts = @('itglue-quick-notes.ps1', 'sync-itglue-organisations-sharepoint.ps1', 'sync-unifi-devices-itglue.ps1', 'test-delegated-tenant-access.ps1')

    function Get-ScriptPath {
        param([Parameter(Mandatory)][string]$Name)
        Join-Path -Path $script:ScriptRoot -ChildPath $Name
    }

    # Stubs so Mock has a command to bind to when the real modules are not installed.
    if (-not (Get-Command -Name 'Get-Secret' -ErrorAction SilentlyContinue)) {
        function Get-Secret { [CmdletBinding()] param($Name, $Vault, [switch]$AsPlainText) }
    }
    function Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param() }
    function Get-OrganizationConfig { [CmdletBinding()] param() }
}

Describe 'automation scripts (mocked)' {
    BeforeEach {
        Mock Get-Secret { 'fake-secret-value' }
        Mock Get-MspAuthHeader { throw 'Get-MspAuthHeader must not be called directly in these tests.' }
        Mock Invoke-MspGraphRequest { }
        Mock Invoke-RestMethod { throw "Unexpected network call to $Uri" }
        Mock Invoke-WebRequest { throw 'Network calls are not allowed in these tests.' }
        Mock Connect-MspExchangeOnline { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MspCustomer {
            [pscustomobject]@{ TenantId = '22222222-2222-4222-8222-222222222222'; DisplayName = 'Contoso'; DefaultDomainName = 'contoso.onmicrosoft.com'; GdapStatus = 'active' }
        }
    }

    Context 'test-delegated-tenant-access.ps1' {
        BeforeEach {
            Mock Test-MspGdapAccess { [pscustomobject]@{ Success = $true; EffectiveRoleNames = @('Global Reader', 'Exchange Administrator'); Steps = @() } }
            Mock Test-MspPartnerAppConsent { [pscustomobject]@{ Success = $false; MissingScopes = @{ 'Microsoft Graph' = @('AuditLog.Read.All') }; Steps = @([pscustomobject]@{ Step = 'Grant Microsoft Graph'; Status = 'Failed'; Detail = 'Missing AuditLog.Read.All' }) } }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = '22222222-2222-4222-8222-222222222222'; displayName = 'Contoso'; verifiedDomains = @([pscustomobject]@{ name = 'contoso.onmicrosoft.com'; isInitial = $true }) } } -ParameterFilter { $Uri -like 'v1.0/organization*' }
            Mock Get-OrganizationConfig { [pscustomobject]@{ DisplayName = 'Contoso' } }
        }

        It 'returns one row per customer with each check' {
            $rows = @(& (Get-ScriptPath 'test-delegated-tenant-access.ps1') -TenantId $script:Tenant -TestExchange)
            $rows.Count | Should -Be 1
            $rows[0].CustomerTenantId | Should -Be $script:CustomerId
            $rows[0].CustomerName | Should -Be 'Contoso'
            $rows[0].GdapCheck | Should -Be 'Passed'
            $rows[0].EffectiveRoles | Should -Be 'Global Reader, Exchange Administrator'
            $rows[0].ConsentCheck | Should -Be 'Failed'
            $rows[0].MissingScopes | Should -Be 'Microsoft Graph: AuditLog.Read.All'
            $rows[0].GraphCheck | Should -Be 'Passed'
            $rows[0].InitialDomain | Should -Be 'contoso.onmicrosoft.com'
            $rows[0].ExchangeCheck | Should -Be 'Passed'
            $rows[0].Status | Should -Be 'Failed'
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }

        It 'flags Graph answering for the wrong tenant' {
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = '33333333-3333-4333-8333-333333333333'; verifiedDomains = @() } } -ParameterFilter { $Uri -like 'v1.0/organization*' }
            $rows = @(& (Get-ScriptPath 'test-delegated-tenant-access.ps1') -TenantId $script:Tenant)
            $rows[0].GraphCheck | Should -Be 'Failed'
        }

        It 'never calls a Graph write' {
            $null = & (Get-ScriptPath 'test-delegated-tenant-access.ps1') -AllCustomers
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -ne 'GET' }
        }
    }

    Context 'sync-itglue-organisations-sharepoint.ps1' {
        BeforeEach {
            Mock Invoke-RestMethod {
                [pscustomobject]@{
                    data  = @(
                        [pscustomobject]@{ id = '101'; attributes = [pscustomobject]@{ name = 'Contoso'; 'short-name' = 'CON' } }
                        [pscustomobject]@{ id = '102'; attributes = [pscustomobject]@{ name = 'Fabrikam'; 'short-name' = 'FAB' } }
                    )
                    links = [pscustomobject]@{ next = $null }
                }
            } -ParameterFilter { $Uri -like 'https://api.itglue.com/organizations*' -and $Method -eq 'GET' }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'list-1'; displayName = 'ITGlue Org Register' } } -ParameterFilter { $PartnerTenant -and $Method -eq 'GET' -and $Uri -like 'v1.0/sites/root/lists?*' }
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{ id = '1'; fields = [pscustomobject]@{ Title = 'Contoso Old'; ShortName = 'CON'; ITGlueID = 101.0; CustomerTenantId = $null } }
                [pscustomobject]@{ id = '9'; fields = [pscustomobject]@{ Title = 'Gone Pty Ltd'; ShortName = 'GONE'; ITGlueID = 999.0; CustomerTenantId = $null } }
            } -ParameterFilter { $PartnerTenant -and $Method -eq 'GET' -and $Uri -like 'v1.0/sites/root/lists/list-1/items*' }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'new' } } -ParameterFilter { $Method -in 'POST', 'PATCH', 'DELETE' }
        }

        It 'reports create, update and stale items without writing' {
            $rows = @(& (Get-ScriptPath 'sync-itglue-organisations-sharepoint.ps1') -ITGlueApiKeySecretName 'ITGlueApiKey' -MatchCustomers)
            $rows.Count | Should -Be 3
            foreach ($row in $rows) {
                $row.PSObject.Properties.Name | Should -Contain 'CustomerTenantId'
                $row.PSObject.Properties.Name | Should -Contain 'CustomerName'
            }
            ($rows | Where-Object ITGlueId -eq '101').Action | Should -Be 'Update (report only)'
            ($rows | Where-Object ITGlueId -eq '101').CustomerTenantId | Should -Be $script:CustomerId
            ($rows | Where-Object ITGlueId -eq '102').Action | Should -Be 'Create (report only)'
            ($rows | Where-Object ITGlueId -eq '999').Action | Should -Be 'Stale (report only)'
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -in 'POST', 'PATCH', 'DELETE' }
        }

        It 'makes no change under -Apply -RemoveStale -WhatIf' {
            $rows = @(& (Get-ScriptPath 'sync-itglue-organisations-sharepoint.ps1') -ITGlueApiKeySecretName 'ITGlueApiKey' -RemoveStale -Apply -WhatIf)
            @($rows | Where-Object Action -eq 'WhatIf').Count | Should -Be 3
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -in 'POST', 'PATCH', 'DELETE' }
        }

        It 'writes through Graph in the partner tenant with -Apply' {
            $null = & (Get-ScriptPath 'sync-itglue-organisations-sharepoint.ps1') -ITGlueApiKeySecretName 'ITGlueApiKey' -RemoveStale -Apply -Confirm:$false
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $PartnerTenant -and $Method -eq 'POST' }
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $PartnerTenant -and $Method -eq 'PATCH' }
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $PartnerTenant -and $Method -eq 'DELETE' }
        }

        It 'refuses to send the API key to a host outside IT Glue' {
            Mock Invoke-RestMethod {
                [pscustomobject]@{ data = @(); links = [pscustomobject]@{ next = 'https://attacker.example/organizations?page=2' } }
            } -ParameterFilter { $Uri -like 'https://api.itglue.com/organizations*' }
            { & (Get-ScriptPath 'sync-itglue-organisations-sharepoint.ps1') -ITGlueApiKeySecretName 'ITGlueApiKey' } | Should -Throw '*outside*'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Uri -like 'https://attacker.example*' }
        }
    }

    Context 'sync-unifi-devices-itglue.ps1' {
        BeforeEach {
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'org-list' } } -ParameterFilter { $Uri -like "*displayName eq 'ITGlue Org Register'*" }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'match-list' } } -ParameterFilter { $Uri -like "*displayName eq 'UniFi - IT Glue match register'*" }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = '5'; fields = [pscustomobject]@{ Title = 'Contoso'; ITGlueID = 101.0; CustomerTenantId = '22222222-2222-4222-8222-222222222222' } } } -ParameterFilter { $Uri -like 'v1.0/sites/root/lists/org-list/items*' }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = '1'; fields = [pscustomobject]@{ UnifiSiteName = 'abc123'; ITGlueLookupId = '5' } } } -ParameterFilter { $Uri -like 'v1.0/sites/root/lists/match-list/items*' }
            Mock Invoke-RestMethod {
                [pscustomobject]@{ offset = 0; limit = 200; count = 2; totalCount = 2; data = @(
                        [pscustomobject]@{ id = 'site-guid-1'; internalReference = 'abc123'; name = 'Contoso HQ' }
                        [pscustomobject]@{ id = 'site-guid-2'; internalReference = 'zzz999'; name = 'Unmatched site' }
                    )
                }
            } -ParameterFilter { $Uri -like 'https://unifi.contoso.com/proxy/network/integration/v1/sites?*' }
            Mock Invoke-RestMethod {
                [pscustomobject]@{ offset = 0; limit = 200; count = 2; totalCount = 2; data = @(
                        [pscustomobject]@{ id = 'd1'; name = 'HQ Gateway'; model = 'UDM Pro'; macAddress = 'aa:bb:cc:dd:ee:01'; ipAddress = '192.0.2.1' }
                        [pscustomobject]@{ id = 'd2'; name = 'HQ AP'; model = 'U7 Pro'; macAddress = 'aa:bb:cc:dd:ee:02'; ipAddress = '192.0.2.20' }
                    )
                }
            } -ParameterFilter { $Uri -like 'https://unifi.contoso.com/proxy/network/integration/v1/sites/site-guid-1/devices*' }
            Mock Invoke-RestMethod {
                [pscustomobject]@{ data = @([pscustomobject]@{ id = '777'; attributes = [pscustomobject]@{ name = 'Old gateway name'; 'primary-ip' = '192.0.2.1'; 'mac-address' = 'AA-BB-CC-DD-EE-01' } }); links = [pscustomobject]@{ next = $null } }
            } -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'https://api.itglue.com/configurations*' }
            Mock Invoke-RestMethod { [pscustomobject]@{ data = [pscustomobject]@{ id = '888' } } } -ParameterFilter { $Method -in 'POST', 'PATCH' }

            $script:UniFiParams = @{
                UniFiBaseUri           = 'https://unifi.contoso.com/proxy/network/integration'
                UniFiApiKeySecretName  = 'UniFiApiKey'
                ITGlueApiKeySecretName = 'ITGlueApiKey'
                ConfigurationTypeId    = 101
                ConfigurationStatusId  = 201
            }
        }

        It 'reports updates, creates and unmatched sites without writing' {
            $rows = @(& (Get-ScriptPath 'sync-unifi-devices-itglue.ps1') @script:UniFiParams -UpdateExisting)
            $rows.Count | Should -Be 3
            ($rows | Where-Object DeviceName -eq 'HQ Gateway').Action | Should -Be 'Update (report only)'
            ($rows | Where-Object DeviceName -eq 'HQ AP').Action | Should -Be 'Create (report only)'
            ($rows | Where-Object Action -eq 'Unmatched').UniFiSite | Should -Match 'Unmatched site'
            ($rows | Where-Object DeviceName -eq 'HQ AP').CustomerTenantId | Should -Be $script:CustomerId
            ($rows | Where-Object DeviceName -eq 'HQ AP').CustomerName | Should -Be 'Contoso'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -in 'POST', 'PATCH' }
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'sync-unifi-devices-itglue.ps1') @script:UniFiParams -UpdateExisting -Apply -WhatIf)
            @($rows | Where-Object Action -eq 'WhatIf').Count | Should -Be 2
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -in 'POST', 'PATCH' }
        }

        It 'leaves existing configurations alone without -UpdateExisting' {
            $rows = @(& (Get-ScriptPath 'sync-unifi-devices-itglue.ps1') @script:UniFiParams -Apply -Confirm:$false)
            ($rows | Where-Object DeviceName -eq 'HQ Gateway').Action | Should -Be 'Differs'
            ($rows | Where-Object DeviceName -eq 'HQ AP').Action | Should -Be 'Created'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -eq 'PATCH' }
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' }
        }

        It 'creates and updates configurations with -Apply -UpdateExisting' {
            $null = & (Get-ScriptPath 'sync-unifi-devices-itglue.ps1') @script:UniFiParams -UpdateExisting -Apply -Confirm:$false
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'https://api.itglue.com/configurations' }
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Uri -eq 'https://api.itglue.com/configurations/777' }
        }

        It 'sends the UniFi key only to the UniFi console' {
            $null = & (Get-ScriptPath 'sync-unifi-devices-itglue.ps1') @script:UniFiParams
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Headers -and (@($Headers.Keys) -ccontains 'X-API-KEY') -and $Uri -notlike 'https://unifi.contoso.com/*' }
        }
    }

    Context 'itglue-quick-notes.ps1' {
        BeforeEach {
            Mock Invoke-RestMethod {
                [pscustomobject]@{
                    data  = @(
                        [pscustomobject]@{ id = '101'; attributes = [pscustomobject]@{ name = 'Contoso'; 'quick-notes' = $null } }
                        [pscustomobject]@{ id = '102'; attributes = [pscustomobject]@{ name = 'Fabrikam'; 'quick-notes' = 'Already documented' } }
                    )
                    links = [pscustomobject]@{ next = $null }
                }
            } -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'https://api.itglue.com/organizations?*' }
            Mock Invoke-RestMethod {
                [pscustomobject]@{ data = [pscustomobject]@{ id = '101'; attributes = [pscustomobject]@{ name = 'Contoso'; 'quick-notes' = 'Existing note' } } }
            } -ParameterFilter { $Method -eq 'GET' -and $Uri -eq 'https://api.itglue.com/organizations/101' }
            Mock Invoke-RestMethod { } -ParameterFilter { $Method -in 'POST', 'PATCH' }
            Mock Get-Secret { 'https://prod-00.australiaeast.logic.azure.com/workflows/example/triggers/manual/paths/invoke' } -ParameterFilter { $Name -eq 'QuickNotesFlowUrl' }
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{ mail = 'tech1@contoso.com'; userPrincipalName = 'tech1@contoso.com'; accountEnabled = $true }
                [pscustomobject]@{ mail = $null; userPrincipalName = 'tech2@contoso.com'; accountEnabled = $false }
            } -ParameterFilter { $PartnerTenant -and $Uri -like 'v1.0/groups/*' }
        }

        It 'picks only organisations without notes and only enabled technicians' {
            $rows = @(& (Get-ScriptPath 'itglue-quick-notes.ps1') -RequestNotes -TechnicianGroupId '00000000-0000-4000-8000-000000000001' -FlowUrlSecretName 'QuickNotesFlowUrl' -ITGlueApiKeySecretName 'ITGlueApiKey')
            $rows.Count | Should -Be 1
            $rows[0].CustomerName | Should -Be 'Contoso'
            $rows[0].Technician | Should -Be 'tech1@contoso.com'
            $rows[0].Action | Should -Be 'ReportOnly'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
        }

        It 'makes no request under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'itglue-quick-notes.ps1') -RequestNotes -Technician 'tech1@contoso.com' -FlowUrlSecretName 'QuickNotesFlowUrl' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
        }

        It 'posts to the flow with -Apply' {
            $null = & (Get-ScriptPath 'itglue-quick-notes.ps1') -RequestNotes -Technician 'tech1@contoso.com' -FlowUrlSecretName 'QuickNotesFlowUrl' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply -Confirm:$false
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -like 'https://prod-00.australiaeast.logic.azure.com/*' }
        }

        It 'HTML encodes the note and makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'itglue-quick-notes.ps1') -AppendNote -OrganizationId 101 -Note '<script>x</script> Call after 2 pm' -Responder 'Jane' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].Detail | Should -Be '&lt;script&gt;x&lt;/script&gt; Call after 2 pm - Jane'
            Should -Invoke Invoke-RestMethod -Times 0 -Exactly -ParameterFilter { $Method -eq 'PATCH' }
        }

        It 'appends to existing notes with -Apply' {
            $null = & (Get-ScriptPath 'itglue-quick-notes.ps1') -AppendNote -OrganizationId 101 -Note 'New note' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply -Confirm:$false
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Body -match 'Existing note<br><br>New note' }
        }
    }

    Context 'Hygiene' {
        It 'contains no retired or unsafe methods outside help text' {
            $pattern = 'Connect-MsolService|Connect-AzureAD|New-AzureAD|AzureRM|New-PSSession|-Authentication\s+Basic|ConvertTo-SecureString|graph\.windows\.net|oauth2/token\b|client_secret|New-PartnerAccessToken|api/login'
            foreach ($file in $script:OwnScripts | ForEach-Object { Get-Item -LiteralPath (Get-ScriptPath $_) }) {
                $code = (Get-Content -LiteralPath $file.FullName -Raw) -replace '(?s)<#.*?#>', ''
                $code | Should -Not -Match $pattern -Because $file.Name
            }
        }

        It 'has the required header and help in every script' {
            foreach ($file in $script:OwnScripts | ForEach-Object { Get-Item -LiteralPath (Get-ScriptPath $_) }) {
                $text = Get-Content -LiteralPath $file.FullName -Raw
                $text | Should -Match '#Requires -Version 7\.4' -Because $file.Name
                $text | Should -Match '#Requires -Modules MspGdap' -Because $file.Name
                $help = Get-Help -Name $file.FullName -Full
                $help.Synopsis | Should -Not -BeNullOrEmpty -Because $file.Name
                @($help.Examples.Example).Count | Should -BeGreaterOrEqual 2 -Because $file.Name
                $text | Should -Match 'Replaces the original \d{4} method:' -Because $file.Name
                $text | Should -Match 'Required GDAP roles:' -Because $file.Name
                $text | Should -Match 'Required partner app permissions:' -Because $file.Name
                $text | Should -Match '\.LINK\s+https://gcit\.com\.au/' -Because $file.Name
                $text | Should -Not -Match '[\u2013\u2014]' -Because "$($file.Name) must not contain en or em dashes"
            }
        }
    }
}
