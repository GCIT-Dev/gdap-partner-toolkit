#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Invoke-MspGraphRequest' {
    BeforeEach {
        Initialize-TestModuleState
        $script:TokenCalls = 0
        $script:GraphCalls = 0
        Mock Start-Sleep -ModuleName MspGdap -MockWith {}
        Mock Get-MspClientCredentialBody -ModuleName MspGdap -MockWith { [ordered]@{ client_assertion_type = 't'; client_assertion = 'a' } }
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { $null }
        Mock Get-Secret -ModuleName MspGdap -MockWith { New-TestSecure 'rt-old-value' }
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith {
            $script:TokenCalls++
            New-TestTokenResponse -TenantId (([uri]"$Uri").Segments[1].TrimEnd('/')) -RefreshToken "rt-new-$($script:TokenCalls)"
        }
    }

    It 'requires -TenantId or -PartnerTenant' {
        $command = Get-Command Invoke-MspGraphRequest
        foreach ($set in $command.ParameterSets) {
            $mandatory = @($set.Parameters | Where-Object IsMandatory | ForEach-Object Name)
            ($mandatory -contains 'TenantId' -or $mandatory -contains 'PartnerTenant') | Should -BeTrue
        }
    }

    It 'follows @odata.nextLink and writes each item' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            if ("$Uri" -like '*skiptoken*') {
                [pscustomobject]@{ value = @([pscustomobject]@{ id = '3' }) }
            }
            else {
                [pscustomobject]@{ value = @([pscustomobject]@{ id = '1' }, [pscustomobject]@{ id = '2' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' }
            }
        }
        $items = @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users')
        $items.id | Should -Be @('1', '2', '3')
        $script:GraphCalls | Should -Be 2
        $script:TokenCalls | Should -Be 1
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { "$Uri" -eq 'https://graph.microsoft.com/v1.0/users' -and $Headers.Authorization -like 'Bearer eyJ*' }
    }

    It 'with -NoPaging writes the items of the first page only and says when more pages exist' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            [pscustomobject]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#users'; value = @([pscustomobject]@{ id = '1' }, [pscustomobject]@{ id = '2' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' }
        }
        $items = @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users' -NoPaging -Verbose 4>&1)
        $verbose = @($items | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] })
        $items = @($items | Where-Object { $_ -isnot [System.Management.Automation.VerboseRecord] })
        $items.id | Should -Be @('1', '2')
        $items | ForEach-Object { $_.PSObject.Properties.Name | Should -Not -Contain '@odata.nextLink' }
        ($verbose.Message -join ' ') | Should -Match 'first page only'
        $script:GraphCalls | Should -Be 1
    }

    It 'with -NoPaging writes a single item of a one-item page and nothing for an empty page' {
        $script:Page = @([pscustomobject]@{ id = 'only' })
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            [pscustomobject]@{ '@odata.context' = 'x'; value = $script:Page }
        }
        $one = @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users?$top=1' -NoPaging)
        $one.Count | Should -Be 1
        $one[0].id | Should -Be 'only'
        $script:Page = @()
        @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users?$top=1' -NoPaging).Count | Should -Be 0
    }

    It 'with -NoPaging returns <Name> unchanged' -ForEach @(
        @{ Name = 'a single entity'; Path = 'organization/x'; Response = { [pscustomobject]@{ id = 'x'; displayName = 'Contoso' } }; Check = { param($r) $r.displayName | Should -Be 'Contoso' } }
        @{ Name = 'report CSV content'; Path = "reports/getMailboxUsageDetail(period='D7')"; Response = { "Report Refresh Date,User Principal Name`nx,a@contoso.com" }; Check = { param($r) $r | Should -BeOfType [string]; $r | Should -Match 'User Principal Name' } }
        @{ Name = '$count text'; Path = 'users/$count'; Response = { '42' }; Check = { param($r) $r | Should -Be '42' } }
        @{ Name = 'an entity whose value property is not an array'; Path = 'x/y'; Response = { [pscustomobject]@{ id = 'p'; value = 'scalar' } }; Check = { param($r) $r.value | Should -Be 'scalar'; $r.id | Should -Be 'p' } }
    ) {
        $script:Response = $Response
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith { & $script:Response }
        $result = @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri $Path -NoPaging)
        $result.Count | Should -Be 1
        & $Check $result[0]
    }

    It 'with -NoPaging returns a write response unchanged' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            [pscustomobject]@{ '@odata.context' = 'x'; value = @('group-1', 'group-2') }
        }
        $result = @(Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users/u1/getMemberGroups' -Method POST -Body @{ securityEnabledOnly = $true } -NoPaging -Confirm:$false)
        $result.Count | Should -Be 1
        $result[0].value | Should -Be @('group-1', 'group-2')
    }

    It 'uses the beta endpoint when asked' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith { [pscustomobject]@{ id = 'x' } }
        $null = Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri '/organization' -ApiVersion beta
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { "$Uri" -eq 'https://graph.microsoft.com/beta/organization' }
    }

    It 'waits for Retry-After on 429 and retries' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            if ($script:GraphCalls -eq 1) { throw (New-TestHttpError -Status 429 -RetryAfterSeconds 7 -Json '{"error":{"code":"TooManyRequests","message":"slow down"}}') }
            [pscustomobject]@{ id = 'org' }
        }
        $result = Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'organization/x'
        $result.id | Should -Be 'org'
        $script:GraphCalls | Should -Be 2
        Should -Invoke Start-Sleep -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 7000 }
    }

    It 'backs off exponentially on 503 without Retry-After' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            if ($script:GraphCalls -le 2) { throw (New-TestHttpError -Status 503) }
            [pscustomobject]@{ id = 'ok' }
        }
        $null = Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'organization/x'
        Should -Invoke Start-Sleep -ModuleName MspGdap -Times 2 -Exactly
    }

    It 'refreshes the token once on 401' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            if ($script:GraphCalls -eq 1) { throw (New-TestHttpError -Status 401 -Json '{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}') }
            [pscustomobject]@{ id = 'ok' }
        }
        $null = Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'organization/x'
        $script:TokenCalls | Should -Be 2
    }

    It 'raises a clear error for 403 without retrying' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            $script:GraphCalls++
            throw (New-TestHttpError -Status 403 -Json '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges to complete the operation.","innerError":{"request-id":"req-1"}}}')
        }
        $caught = $null
        try { Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users' } catch { $caught = $_ }
        $caught.FullyQualifiedErrorId | Should -BeLike 'MspGdap.Graph.Http403*'
        $caught.Exception.Message | Should -BeLike '*Authorization_RequestDenied*Insufficient privileges*req-1*'
        $caught.Exception.Message | Should -Not -Match 'Bearer|eyJ'
        $script:GraphCalls | Should -Be 1
    }

    It 'does not call Graph for a write under -WhatIf' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith { $script:GraphCalls++ }
        Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users/abc' -Method PATCH -Body @{ usageLocation = 'AU' } -WhatIf
        $script:GraphCalls | Should -Be 0
    }

    It 'sends a JSON body on writes' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith { $null }
        Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users/abc' -Method PATCH -Body @{ usageLocation = 'AU' } -Confirm:$false
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Body -eq '{"usageLocation":"AU"}' -and $ContentType -eq 'application/json' }
    }

    It 'refuses to send the token to another host' -ForEach @(
        @{ Uri = 'https://graph.microsoft.com.evil.example/v1.0/users' }
        @{ Uri = 'http://graph.microsoft.com/v1.0/users' }
        @{ Uri = 'https://example.com/v1.0/users' }
    ) {
        { Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri $Uri } | Should -Throw -ErrorId 'MspGdap.Graph.ForeignHost*'
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'refuses a nextLink that points at another host' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like 'https://graph.microsoft.com/*' } -MockWith {
            [pscustomobject]@{ value = @([pscustomobject]@{ id = '1' }); '@odata.nextLink' = 'https://attacker.example/v1.0/users?$skiptoken=1' }
        }
        { $null = Invoke-MspGraphRequest -TenantId $script:TestCustomerTenantId -Uri 'users' } | Should -Throw -ErrorId 'MspGdap.Graph.ForeignHost*'
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 0 -Exactly -ParameterFilter { "$Uri" -like 'https://attacker.example/*' }
    }

    It 'refuses the partner tenant ID as -TenantId' {
        { Invoke-MspGraphRequest -TenantId $script:TestPartnerTenantId -Uri 'users' } | Should -Throw -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed*'
    }
}

Describe 'Invoke-MspGraphBatch' {
    BeforeEach {
        Initialize-TestModuleState
        $script:BatchCalls = 0
        Mock Start-Sleep -ModuleName MspGdap -MockWith {}
        Mock Get-MspAuthHeader -ModuleName MspGdap -MockWith { @{ Authorization = 'Bearer test' } }
    }

    It 'returns one result per request and retries throttled sub-requests' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -eq 'https://graph.microsoft.com/v1.0/$batch' } -MockWith {
            $script:BatchCalls++
            $requests = ($Body | ConvertFrom-Json).requests
            $responses = foreach ($r in $requests) {
                if ($r.id -eq '2' -and $script:BatchCalls -eq 1) {
                    [pscustomobject]@{ id = $r.id; status = 429; headers = [pscustomobject]@{ 'Retry-After' = '3' } }
                }
                else {
                    [pscustomobject]@{ id = $r.id; status = 200; body = [pscustomobject]@{ id = "obj-$($r.id)" } }
                }
            }
            [pscustomobject]@{ responses = @($responses) }
        }
        $results = @(Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(
                @{ method = 'GET'; url = '/users/a' }
                @{ method = 'GET'; url = 'users/b' }
            ))
        $results.Count | Should -Be 2
        @($results | Where-Object Success).Count | Should -Be 2
        $script:BatchCalls | Should -Be 2
        Should -Invoke Start-Sleep -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { $Seconds -eq 3 }
    }

    It 'chunks requests into batches of 20' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith {
            $script:BatchCalls++
            $requests = ($Body | ConvertFrom-Json).requests
            [pscustomobject]@{ responses = @($requests | ForEach-Object { [pscustomobject]@{ id = $_.id; status = 200 } }) }
        }
        $requests = 1..45 | ForEach-Object { @{ method = 'GET'; url = "/users/$_" } }
        $results = @(Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request $requests)
        $results.Count | Should -Be 45
        $script:BatchCalls | Should -Be 3
    }

    It 'honours -WhatIf when the batch contains writes' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { $script:BatchCalls++ }
        Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(@{ method = 'DELETE'; url = '/users/a' }) -WhatIf
        $script:BatchCalls | Should -Be 0
    }
}

Describe 'Resolve-MspTenantId' {
    BeforeEach {
        Initialize-TestModuleState
    }

    It 'returns a GUID unchanged in lower case without a network call' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { throw 'should not be called' }
        Resolve-MspTenantId -Tenant $script:TestCustomerTenantId.ToUpper() | Should -Be $script:TestCustomerTenantId
    }

    It 'resolves a domain through the OpenID discovery document and caches it' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith {
            [pscustomobject]@{ issuer = "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/v2.0" }
        }
        Resolve-MspTenantId -Tenant 'Contoso.onmicrosoft.com' | Should -Be $script:TestCustomerTenantId
        Resolve-MspTenantId -Tenant 'contoso.onmicrosoft.com' | Should -Be $script:TestCustomerTenantId
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { "$Uri" -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com/v2.0/.well-known/openid-configuration' }
    }

    It 'rejects values that are not domains' -ForEach @(
        @{ Value = 'contoso.com/../evil' }
        @{ Value = 'not a domain' }
        @{ Value = 'common' }
        @{ Value = 'https://contoso.com' }
    ) {
        { Resolve-MspTenantId -Tenant $Value } | Should -Throw -ErrorId 'MspGdap.Tenant.InvalidIdentifier*'
    }

    It 'reports an unknown domain' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { throw (New-TestHttpError -Status 400 -Json '{"error":"invalid_tenant"}') }
        { Resolve-MspTenantId -Tenant 'no-such-tenant.example' } | Should -Throw -ErrorId 'MspGdap.Tenant.NotFound*'
    }
}

Describe 'Get-MspCustomer' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Invoke-MspGraphRequest -ModuleName MspGdap -ParameterFilter { $Uri -like 'contracts*' } -MockWith {
            [pscustomobject]@{ customerId = $script:TestCustomerTenantId; displayName = 'Contoso'; defaultDomainName = 'contoso.onmicrosoft.com' }
        }
        Mock Invoke-MspGraphRequest -ModuleName MspGdap -ParameterFilter { $Uri -like 'tenantRelationships/*' } -MockWith {
            [pscustomobject]@{
                id = 'rel-1'; displayName = 'contoso-gdap'; status = 'active'; endDateTime = '2028-01-01T00:00:00Z'; autoExtendDuration = 'P180D'
                customer = [pscustomobject]@{ tenantId = $script:TestCustomerTenantId; displayName = 'Contoso' }
                accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = '29232cdf-9323-42fd-ade2-1d097af3e4de' }) }
            }
            [pscustomobject]@{
                id = 'rel-2'; displayName = 'fabrikam-gdap'; status = 'approvalPending'
                customer = [pscustomobject]@{ tenantId = $script:TestOtherTenantId; displayName = 'Fabrikam' }
                accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = '62e90394-69f5-4237-9190-012177145e10' }) }
            }
        }
    }

    It 'lists contract customers from the partner tenant' {
        $customers = @(Get-MspCustomer)
        $customers.Count | Should -Be 1
        $customers[0].TenantId | Should -Be $script:TestCustomerTenantId
        Should -Invoke Invoke-MspGraphRequest -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { $PartnerTenant -and $Uri -like 'contracts*' }
        Should -Invoke Invoke-MspGraphRequest -ModuleName MspGdap -Times 0 -Exactly -ParameterFilter { $Uri -like 'tenantRelationships/*' }
    }

    It 'adds GDAP status and includes GDAP-only customers' {
        $customers = @(Get-MspCustomer -IncludeGdapStatus)
        $customers.Count | Should -Be 2
        $contoso = $customers | Where-Object TenantId -eq $script:TestCustomerTenantId
        $contoso.GdapStatus | Should -Be 'active'
        $contoso.ActiveRelationshipCount | Should -Be 1
        $contoso.IncludesGlobalAdministrator | Should -BeFalse
        $fabrikam = $customers | Where-Object TenantId -eq $script:TestOtherTenantId
        $fabrikam.HasContract | Should -BeFalse
        $fabrikam.GdapStatus | Should -Be 'approvalPending'
        $fabrikam.PendingApprovalCount | Should -Be 1
        $fabrikam.Relationships[0].IncludesGlobalAdministrator | Should -BeTrue
    }

    It 'filters by name' {
        @(Get-MspCustomer -Name 'Fab*' -IncludeGdapStatus).DisplayName | Should -Be 'Fabrikam'
    }
}
