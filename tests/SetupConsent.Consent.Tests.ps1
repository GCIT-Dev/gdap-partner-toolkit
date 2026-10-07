#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
    $customer = $global:MspTest.CustomerTenantId
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $exoAppId = '00000002-0000-0ff1-ce00-000000000000'
    $pcAppId = 'fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd'

    # The error Microsoft Entra ID returns when the partner app asks for a token in a customer it is not consented in.
    function New-ConsentMissingError {
        param([string]$Code = 'AADSTS65001')
        New-Object System.Exception ("{0}: The user or administrator has not consented to use the application with ID '{1}' named 'Contoso MSP Worker'. Send an interactive authorization request for this user and resource." -f $Code, $global:MspTest.PartnerAppId)
    }

    # A consent state in which every desired scope is present, as Graph shows it once consent has replicated.
    function New-FullConsentState {
        param([string]$TenantId, [object[]]$Grant)
        $resources = foreach ($g in $Grant) {
            [pscustomobject]@{ ResourceAppId = $g.ResourceAppId; ResourceDisplayName = $g.ResourceDisplayName; ResourcePresent = $true; ResourceSpId = "sp-$($g.ResourceAppId)"; GrantId = "grant-$($g.ResourceAppId)"; DesiredScopes = @($g.Scopes); CurrentScopes = @($g.Scopes); MissingScopes = @(); ExtraScopes = @() }
        }
        [pscustomobject]@{ TenantId = $TenantId; AppServicePrincipal = [pscustomobject]@{ id = 'app-sp' }; Resources = @($resources); OtherGrants = @(); AppRoleAssignments = @() }
    }

    function Reset-ConsentFake {
        param([hashtable]$Grants = @{}, [bool]$AppSpPresent = $true, [bool]$ExoPresent = $true, [string]$PostBehaviour = 'apply', [bool]$Relationship = $true, [bool]$GraphFails = $false, [bool]$NotConsented = $false)
        $global:MspTest.State = @{ Grants = $Grants; AppSpPresent = $AppSpPresent; ExoPresent = $ExoPresent; PostBehaviour = $PostBehaviour; GraphFails = $GraphFails; NotConsented = $NotConsented }
        $relationships = if ($Relationship) { @([pscustomobject]@{ id = 'rel'; displayName = 'MspGdap-1'; status = 'active'; customer = [pscustomobject]@{ tenantId = $customer; displayName = 'Fabrikam' }; accessDetails = [pscustomobject]@{ unifiedRoles = @() } }) } else { @() }
        $global:MspTest.State.Relationships = $relationships
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern 'delegatedAdminRelationships$' -Response { $global:MspTest.State.Relationships }
            New-FakeRoute -Pattern "applications\?\`$filter=appId eq '22222222" -Response {
                [pscustomobject]@{ id = 'app-obj'; appId = $global:MspTest.PartnerAppId; displayName = 'Contoso MSP Worker'; requiredResourceAccess = @(
                        [pscustomobject]@{ resourceAppId = $graphAppId; resourceAccess = @([pscustomobject]@{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; type = 'Scope' }, [pscustomobject]@{ id = '06da0dbc-49e2-44d2-8312-53f166ab848a'; type = 'Scope' }, [pscustomobject]@{ id = 'df021288-bdef-4463-88db-98f22de89214'; type = 'Role' }) },
                        [pscustomobject]@{ resourceAppId = $exoAppId; resourceAccess = @([pscustomobject]@{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; type = 'Scope' }) },
                        [pscustomobject]@{ resourceAppId = $pcAppId; resourceAccess = @([pscustomobject]@{ id = '1cebfa2a-fb4d-419e-b5f9-839b4383e05a'; type = 'Scope' }) }
                    ) }
            }
            New-FakeRoute -Pattern 'servicePrincipals\?\$filter=appId eq' -Response {
                param($Body, $Uri, $TenantId)
                if ($global:MspTest.State.GraphFails -and $TenantId) { throw (New-FakeHttpError -StatusCode 403 -Message 'Forbidden') }
                # Before the first consent the partner app cannot get a token in the customer at all.
                if ($global:MspTest.State.NotConsented -and $TenantId) { throw (New-ConsentMissingError) }
                $appId = [regex]::Match($Uri, "appId eq '([^']+)'").Groups[1].Value
                if (-not $TenantId) {
                    switch ($appId) {
                        $graphAppId { [pscustomobject]@{ id = 'p-graph'; appId = $appId; displayName = 'Microsoft Graph'; oauth2PermissionScopes = @([pscustomobject]@{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; value = 'User.Read' }, [pscustomobject]@{ id = '06da0dbc-49e2-44d2-8312-53f166ab848a'; value = 'Directory.Read.All' }) } }
                        $exoAppId { [pscustomobject]@{ id = 'p-exo'; appId = $appId; displayName = 'Office 365 Exchange Online'; oauth2PermissionScopes = @([pscustomobject]@{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; value = 'Exchange.Manage' }) } }
                    }
                    return
                }
                switch ($appId) {
                    $global:MspTest.PartnerAppId { if ($global:MspTest.State.AppSpPresent) { [pscustomobject]@{ id = 'app-sp'; appId = $appId; displayName = 'Contoso MSP Worker' } } }
                    $graphAppId { [pscustomobject]@{ id = 'graph-sp'; appId = $appId; displayName = 'Microsoft Graph' } }
                    'c5393580-f805-4401-95e8-94b7a6ef2fc2' { [pscustomobject]@{ id = 'mgmt-sp'; appId = $appId; displayName = 'Office 365 Management APIs' } }
                    $exoAppId { if ($global:MspTest.State.ExoPresent) { [pscustomobject]@{ id = 'exo-sp'; appId = $appId; displayName = 'Office 365 Exchange Online' } } }
                }
            }
            New-FakeRoute -Pattern 'oauth2PermissionGrants\?' -Response {
                foreach ($key in $global:MspTest.State.Grants.Keys) { [pscustomobject]@{ id = "grant-$key"; clientId = 'app-sp'; consentType = 'AllPrincipals'; resourceId = $key; scope = $global:MspTest.State.Grants[$key] } }
            }
            New-FakeRoute -Pattern 'servicePrincipals/app-sp/appRoleAssignments' -Response { @() }
            New-FakeRoute -Pattern 'servicePrincipals/mgmt-sp\?' -Response { [pscustomobject]@{ id = 'mgmt-sp'; appId = 'c5393580-f805-4401-95e8-94b7a6ef2fc2'; displayName = 'Office 365 Management APIs' } }
        )
        Mock Invoke-MspPartnerCenterRequest {
            if ($Method -eq 'DELETE') {
                $global:MspTest.State.Grants = @{}
                return [pscustomobject]@{ StatusCode = 204; Success = $true; IsMfaCompliant = $true; MfaRequired = $false; Error = $null }
            }
            $grant = $Body.applicationGrants[0]
            $global:MspTest.State.Posts = @($global:MspTest.State.Posts) + $grant.enterpriseApplicationId
            switch ($global:MspTest.State.PostBehaviour) {
                'mfa' { return [pscustomobject]@{ StatusCode = 401; Success = $false; IsMfaCompliant = $false; MfaRequired = $true; Error = 'MFA required' } }
                'noeffect' { return [pscustomobject]@{ StatusCode = 201; Success = $true; IsMfaCompliant = $true; MfaRequired = $false; Error = $null } }
                default {
                    $spId = switch ($grant.enterpriseApplicationId) { $graphAppId { 'graph-sp' } 'c5393580-f805-4401-95e8-94b7a6ef2fc2' { 'mgmt-sp' } default { 'exo-sp' } }
                    $global:MspTest.State.AppSpPresent = $true
                    $global:MspTest.State.NotConsented = $false
                    $global:MspTest.State.Grants[$spId] = ($grant.scope -split ', ') -join ' '
                    return [pscustomobject]@{ StatusCode = 201; Success = $true; IsMfaCompliant = $true; MfaRequired = $false; Error = $null }
                }
            }
        }
    }
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'Grant-MspPartnerAppConsent' {
    It 'posts nothing with -WhatIf' {
        Reset-ConsentFake -AppSpPresent $false
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -WhatIf
        $r.Outcome | Should -Be 'WhatIf'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'posts one resource per call, scopes joined with comma and space, never Partner Center, then confirms through Graph' {
        Reset-ConsentFake -AppSpPresent $false
        $r = Grant-MspPartnerAppConsent -CustomerTenantId 'fabrikam.onmicrosoft.com' -Confirm:$false
        $r.Success | Should -BeTrue
        @($r.Steps | Where-Object Status -eq 'Changed').Count | Should -Be 2
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'POST' -and @($Body.applicationGrants).Count -eq 1 }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Body.applicationGrants[0].scope -eq 'Directory.Read.All, User.Read' -and $Path -eq "customers/$customer/applicationconsents" -and $Body.displayName -eq 'Contoso MSP Worker' }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0 -ParameterFilter { $Body.applicationGrants[0].enterpriseApplicationId -eq $pcAppId }
    }
    It 'is idempotent: no POST when every scope is already consented' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All'; 'exo-sp' = 'Exchange.Manage' }
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false
        $r.Success | Should -BeTrue
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'does not report a pass when Partner Center accepts but Graph does not show the grant' {
        Reset-ConsentFake -AppSpPresent $false -PostBehaviour 'noeffect'
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false -ReadbackTimeoutSeconds 10
        $r.Success | Should -BeFalse
        @($r.Steps | Where-Object { $_.Step -like 'Consent:*' -and $_.Status -eq 'Changed' }).Count | Should -Be 0
    }
    It 'refuses to top up scopes without -Force and never deletes' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read'; 'exo-sp' = 'Exchange.Manage' }
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Consent: Microsoft Graph').Detail | Should -Match 'Directory.Read.All'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'with -Force deletes then re-posts every resource and confirms' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read'; 'exo-sp' = 'Exchange.Manage' }
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Force -Confirm:$false
        $r.Success | Should -BeTrue
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Path -eq "customers/$customer/applicationconsents/$($global:MspTest.PartnerAppId)" }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'POST' }
    }
    It 'stops every customer after Partner Center says MFA required' {
        Reset-ConsentFake -AppSpPresent $false -PostBehaviour 'mfa'
        $results = @(Grant-MspPartnerAppConsent -CustomerTenantId $customer, $customer -Confirm:$false)
        $results.Count | Should -Be 2
        $results.Success | Should -Not -Contain $true
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly
        ($results[1].Steps | Select-Object -Last 1).Detail | Should -Match 'MFA'
    }
    It 'refuses a customer without an active GDAP relationship' {
        Reset-ConsentFake -AppSpPresent $false -Relationship $false
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false
        $r.Success | Should -BeFalse
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'fails the resource whose service principal is missing in the customer' {
        Reset-ConsentFake -AppSpPresent $false -ExoPresent $false
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Consent: Office 365 Exchange Online').Status | Should -Be 'Failed'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly
    }
    It 'reports Unknown and posts nothing when the customer cannot be read' {
        Reset-ConsentFake -GraphFails $true
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false
        $r.Outcome | Should -Be 'Failed'
        ($r.Steps | Where-Object Step -eq 'Read current consent').Status | Should -Be 'Unknown'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'treats AADSTS65001 on the first read as not consented yet, posts every resource and confirms with a second read' {
        Reset-ConsentFake -AppSpPresent $false
        $global:MspTest.State.Reads = 0
        Mock Get-MspCustomerConsentState {
            $global:MspTest.State.Reads++
            if ($global:MspTest.State.Reads -eq 1) { throw (New-ConsentMissingError) }
            New-FullConsentState -TenantId $TenantId -Grant $Grant
        }
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Confirm:$false
        $r.Success | Should -BeTrue
        $read = $r.Steps | Where-Object Step -eq 'Read current consent'
        $read.Status | Should -Be 'Passed'
        $read.Detail | Should -Match 'Not consented yet \(AADSTS65001\)'
        @($r.Steps | Where-Object { $_.Step -like 'Consent:*' -and $_.Status -eq 'Changed' }).Step | Sort-Object | Should -Be @('Consent: Microsoft Graph', 'Consent: Office 365 Exchange Online')
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'POST' }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Body.applicationGrants[0].enterpriseApplicationId -eq $graphAppId -and $Body.applicationGrants[0].scope -eq 'Directory.Read.All, User.Read' }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Body.applicationGrants[0].enterpriseApplicationId -eq $exoAppId -and $Body.applicationGrants[0].scope -eq 'Exchange.Manage' }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0 -ParameterFilter { $Method -eq 'DELETE' }
        Should -Invoke Get-MspCustomerConsentState -Times 2 -Exactly
    }
    It 'does not report Changed when the confirming read still fails after a first-time consent' {
        Reset-ConsentFake -AppSpPresent $false
        Mock Get-MspCustomerConsentState { throw (New-ConsentMissingError) }
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Confirm:$false -ReadbackTimeoutSeconds 5 -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'POST' }
        @($r.Steps | Where-Object { $_.Step -like 'Consent:*' }).Status | Should -Be @('Unknown', 'Unknown')
        @($r.Steps | Where-Object Status -eq 'Changed').Count | Should -Be 0
    }
    It 'consents a never-consented customer end to end when Graph refuses the app until consent exists' {
        Reset-ConsentFake -AppSpPresent $false -NotConsented $true
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Confirm:$false
        $r.Success | Should -BeTrue
        ($r.Steps | Where-Object Step -eq 'Read current consent').Detail | Should -Match 'AADSTS65001'
        @($r.Steps | Where-Object Status -eq 'Changed').Count | Should -Be 2
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'POST' }
        $global:MspTest.State.Grants['graph-sp'] | Should -Be 'Directory.Read.All User.Read'
        $global:MspTest.State.Grants['exo-sp'] | Should -Be 'Exchange.Manage'
    }
    It 'still reports Unknown and posts nothing for a read error that is not AADSTS65001 (<Name>)' -ForEach @(
        @{ Name = 'AADSTS650052'; MakeError = { New-ConsentMissingError -Code 'AADSTS650052' } }
        @{ Name = 'AADSTS50020'; MakeError = { New-Object System.Exception 'AADSTS50020: User account from identity provider does not exist in tenant.' } }
        @{ Name = 'HTTP 403'; MakeError = { New-FakeHttpError -StatusCode 403 -Message 'Forbidden' } }
    ) {
        Reset-ConsentFake -AppSpPresent $false
        $global:MspTest.State.ReadError = $MakeError
        Mock Get-MspCustomerConsentState { throw (& $global:MspTest.State.ReadError) }
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        $read = $r.Steps | Where-Object Step -eq 'Read current consent'
        $read.Status | Should -Be 'Unknown'
        $read.Detail | Should -Match 'nothing was posted'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'honours -ExcludeResource' {
        Reset-ConsentFake -AppSpPresent $false
        $null = Grant-MspPartnerAppConsent -CustomerTenantId $customer -ExcludeResource 'Office 365 Exchange*' -Confirm:$false
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly
    }
    It 'refuses an -AppId that is not the configured partner app, before any call' {
        Reset-ConsentFake -AppSpPresent $false
        { Grant-MspPartnerAppConsent -TenantId $customer -AppId '55555555-5555-5555-5555-555555555555' -Confirm:$false } | Should -Throw -ErrorId 'MspGdap.Consent.AppIdMismatch*'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'accepts the configured app ID passed explicitly' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All'; 'exo-sp' = 'Exchange.Manage' }
        (Grant-MspPartnerAppConsent -TenantId $customer -AppId $global:MspTest.PartnerAppId -Confirm:$false).Success | Should -BeTrue
    }
    It 'refuses the partner tenant as a target' {
        Reset-ConsentFake
        $r = Grant-MspPartnerAppConsent -TenantId $global:MspTest.PartnerTenantId -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'partner tenant'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'still accepts -CustomerTenantId as an alias' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All'; 'exo-sp' = 'Exchange.Manage' }
        (Grant-MspPartnerAppConsent -CustomerTenantId $customer -Confirm:$false).Success | Should -BeTrue
    }
    It 'with -Force keeps extra scopes and grants for other resources when it replaces the consent' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Mail.Send'; 'exo-sp' = 'Exchange.Manage'; 'mgmt-sp' = 'ActivityFeed.Read' }
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Force -Confirm:$false
        $r.Success | Should -BeTrue
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Body.applicationGrants[0].enterpriseApplicationId -eq $graphAppId -and ($Body.applicationGrants[0].scope -split ', ') -contains 'Mail.Send' -and ($Body.applicationGrants[0].scope -split ', ') -contains 'Directory.Read.All' }
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Body.applicationGrants[0].enterpriseApplicationId -eq 'c5393580-f805-4401-95e8-94b7a6ef2fc2' -and $Body.applicationGrants[0].scope -eq 'ActivityFeed.Read' }
        $global:MspTest.State.Grants['mgmt-sp'] | Should -Be 'ActivityFeed.Read'
    }
    It 'with -Force -WhatIf lists everything it would re-post and deletes nothing' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read'; 'exo-sp' = 'Exchange.Manage' }
        $r = Grant-MspPartnerAppConsent -TenantId $customer -Force -WhatIf
        $r.Outcome | Should -Be 'WhatIf'
        ($r.Steps | Where-Object Step -eq 'Consent: Microsoft Graph').Detail | Should -Match 'Would re-post'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'refuses a manifest whose value does not match the published scope id' {
        Reset-ConsentFake -AppSpPresent $false
        $path = Join-Path $TestDrive 'bad.json'
        @{ requiredResourceAccess = @(@{ resourceAppId = $exoAppId; resourceDisplayName = 'Office 365 Exchange Online'; resourceAccess = @(@{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; type = 'Scope'; value = 'Directory.ReadWrite.All' }) }) } | ConvertTo-Json -Depth 6 | Set-Content -Path $path
        { Grant-MspPartnerAppConsent -TenantId $customer -ManifestPath $path -Confirm:$false } | Should -Throw '*manifest says*'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'refuses a manifest scope the app registration does not request, unless -Force' {
        Reset-ConsentFake -AppSpPresent $false
        $path = Join-Path $TestDrive 'extra.json'
        @{ requiredResourceAccess = @(@{ resourceAppId = $graphAppId; resourceDisplayName = 'Microsoft Graph'; resourceAccess = @(@{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; type = 'Scope'; value = 'User.Read' }, @{ id = '06da0dbc-49e2-44d2-8312-53f166ab848a'; type = 'Scope'; value = 'Directory.Read.All' }, @{ id = 'aaaaaaaa-0000-0000-0000-000000000000'; type = 'Scope'; value = 'Not.Published' }) }) } | ConvertTo-Json -Depth 6 | Set-Content -Path $path
        { Grant-MspPartnerAppConsent -TenantId $customer -ManifestPath $path -Confirm:$false } | Should -Throw '*not a delegated scope*'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'targets every customer from Get-MspCustomer with -AllCustomers' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All'; 'exo-sp' = 'Exchange.Manage' }
        Mock Get-MspCustomer { [pscustomobject]@{ TenantId = $customer; DisplayName = 'Fabrikam' }; [pscustomobject]@{ TenantId = $null; DisplayName = 'No tenant' } }
        $results = @(Grant-MspPartnerAppConsent -AllCustomers -Confirm:$false)
        $results.Count | Should -Be 1
        $results[0].TenantId | Should -Be $customer
        $results[0].CustomerName | Should -Be 'Fabrikam'
    }
    It 'normalises customer objects from Partner Center, GDAP and plain strings' {
        $refs = @(
            [pscustomobject]@{ id = 'pc-1'; companyProfile = [pscustomobject]@{ tenantId = 'aaaaaaaa-0000-0000-0000-000000000001'; companyName = 'Contoso' } }
            [pscustomobject]@{ customer = [pscustomobject]@{ tenantId = 'bbbbbbbb-0000-0000-0000-000000000002'; displayName = 'Fabrikam' } }
            'cccccccc-0000-0000-0000-000000000003'
            $null
        ) | ConvertTo-MspCustomerReference
        $refs.Count | Should -Be 3
        $refs[0].TenantId | Should -Be 'aaaaaaaa-0000-0000-0000-000000000001'
        $refs[0].PartnerCenterId | Should -Be 'pc-1'
        $refs[1].DisplayName | Should -Be 'Fabrikam'
        $refs[2].TenantId | Should -Be 'cccccccc-0000-0000-0000-000000000003'
    }
    It 'takes scopes from a manifest when -ManifestPath is given' {
        Reset-ConsentFake -AppSpPresent $false
        $path = Join-Path $TestDrive 'm.json'
        @{ requiredResourceAccess = @(@{ resourceAppId = $exoAppId; resourceDisplayName = 'Office 365 Exchange Online'; resourceAccess = @(@{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; type = 'Scope'; value = 'Exchange.Manage' }) }) } | ConvertTo-Json -Depth 6 | Set-Content -Path $path
        $r = Grant-MspPartnerAppConsent -CustomerTenantId $customer -ManifestPath $path -Confirm:$false
        $r.Success | Should -BeTrue
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Body.applicationGrants[0].scope -eq 'Exchange.Manage' }
    }
}

Describe 'Remove-MspPartnerAppConsent' {
    It 'deletes through Partner Center and confirms through Graph' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        $r = Remove-MspPartnerAppConsent -TenantId $customer -Confirm:$false
        $r.Success | Should -BeTrue
        ($r.Steps | Where-Object Step -eq 'Confirm through Graph').Status | Should -Be 'Changed'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Path -eq "customers/$customer/applicationconsents/$($global:MspTest.PartnerAppId)" }
    }
    It 'deletes nothing with -WhatIf' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        (Remove-MspPartnerAppConsent -TenantId $customer -WhatIf).Outcome | Should -Be 'WhatIf'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'fails when grants remain after the DELETE' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        Mock Invoke-MspPartnerCenterRequest { [pscustomobject]@{ StatusCode = 404; Success = $false; IsMfaCompliant = $true; MfaRequired = $false; Error = 'Not found' } }
        $r = Remove-MspPartnerAppConsent -TenantId $customer -Confirm:$false -ReadbackTimeoutSeconds 5 -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Confirm through Graph').Detail | Should -Match 'Enterprise apps'
    }
    It 'stops after MFA required' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        Mock Invoke-MspPartnerCenterRequest { [pscustomobject]@{ StatusCode = 401; Success = $false; IsMfaCompliant = $false; MfaRequired = $true; Error = 'MFA required' } }
        $results = @(Remove-MspPartnerAppConsent -TenantId $customer, $customer -Confirm:$false -ErrorAction SilentlyContinue)
        $results.Count | Should -Be 2
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 1 -Exactly
    }
    It 'refuses another app ID' {
        { Remove-MspPartnerAppConsent -TenantId $customer -AppId '55555555-5555-5555-5555-555555555555' -Confirm:$false } | Should -Throw -ErrorId 'MspGdap.Consent.AppIdMismatch*'
    }
}

Describe 'Test-MspPartnerAppConsent' {
    It 'passes when every scope is present' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All'; 'exo-sp' = 'Exchange.Manage' }
        (Test-MspPartnerAppConsent -TenantId $customer).Success | Should -BeTrue
    }
    It 'fails and lists missing scopes' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        $r = Test-MspPartnerAppConsent -TenantId $customer
        $r.Success | Should -BeFalse
        $r.MissingScopes['Microsoft Graph'] | Should -Be @('Directory.Read.All')
        $r.MissingScopes['Office 365 Exchange Online'] | Should -Be @('Exchange.Manage')
    }
    It 'fails when the partner app has no service principal' {
        Reset-ConsentFake -AppSpPresent $false
        (Test-MspPartnerAppConsent -TenantId $customer).Success | Should -BeFalse
    }
    It 'warns about extra scopes without failing' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read Directory.Read.All Mail.Send'; 'exo-sp' = 'Exchange.Manage' }
        $r = Test-MspPartnerAppConsent -TenantId $customer
        $r.Success | Should -BeTrue
        ($r.Steps | Where-Object Step -like 'Extra scopes*').Status | Should -Be 'Warning'
    }
    It 'reports Failed "Not consented" on AADSTS65001' {
        Reset-ConsentFake -AppSpPresent $false
        Mock Get-MspCustomerConsentState { throw (New-ConsentMissingError) }
        $r = Test-MspPartnerAppConsent -TenantId $customer
        $r.Success | Should -BeFalse
        $r.Outcome | Should -Be 'Failed'
        $step = $r.Steps | Where-Object Step -eq 'Read consent'
        $step.Status | Should -Be 'Failed'
        $step.Detail | Should -Match '^Not consented \(AADSTS65001\)'
        $step.Detail | Should -Match 'Grant-MspPartnerAppConsent'
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
    It 'reports Failed "Not consented" when Graph itself refuses the never-consented app' {
        Reset-ConsentFake -AppSpPresent $false -NotConsented $true
        $r = Test-MspPartnerAppConsent -TenantId $customer
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Read consent').Status | Should -Be 'Failed'
    }
    It 'reports Unknown, not Failed, for other read errors' {
        Reset-ConsentFake -GraphFails $true
        $r = Test-MspPartnerAppConsent -TenantId $customer
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Read consent').Status | Should -Be 'Unknown'
    }
    It 'never writes' {
        Reset-ConsentFake -Grants @{ 'graph-sp' = 'User.Read' }
        $null = Test-MspPartnerAppConsent -TenantId $customer
        @(Get-FakeGraphCall | Where-Object Method -ne 'GET').Count | Should -Be 0
        Should -Invoke Invoke-MspPartnerCenterRequest -Times 0
    }
}
