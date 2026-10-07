#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
    $customer = $global:MspTest.CustomerTenantId
    $globalAdmin = '62e90394-69f5-4237-9190-012177145e10'
    $exchangeAdmin = '29232cdf-9323-42fd-ade2-1d097af3e4de'
    $helpdesk = '729827e3-9c14-49f7-bb1b-9608f156bbb8'
    $readers = '88d8e3e3-8f55-4a1e-953a-9b9898b8876b'
    function New-Relationship {
        param([string]$Id = 'rel-1', [string]$Status = 'active', [string[]]$Roles = @($exchangeAdmin, $helpdesk, $readers), [string]$Tenant = $customer, [string]$Name = 'MspGdap-1', [datetime]$Created = [datetime]::UtcNow.AddDays(-10))
        [pscustomobject]@{
            id = $Id; displayName = $Name; status = $Status; duration = 'P730D'; autoExtendDuration = 'P180D'
            createdDateTime = $Created; endDateTime = [datetime]::UtcNow.AddDays(700)
            customer = [pscustomobject]@{ tenantId = $Tenant; displayName = 'Fabrikam' }
            accessDetails = [pscustomobject]@{ unifiedRoles = @($Roles | ForEach-Object { [pscustomobject]@{ roleDefinitionId = $_ } }) }
        }
    }
    $mapPath = Join-Path $TestDrive 'map.json'
    @{ accessMapVersion = 1; assignments = @(
            @{ groupDisplayName = 'GDAP-Helpdesk'; groupId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; roles = @('Helpdesk Administrator', 'Directory Readers') },
            @{ groupDisplayName = 'GDAP-Exchange'; groupId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; roles = @('Exchange Administrator') }
        ) } | ConvertTo-Json -Depth 6 | Set-Content -Path $mapPath
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'Get-MspGdapRelationship' {
    BeforeEach {
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern 'delegatedAdminRelationships$' -Response {
                @(
                    (New-Relationship -Id 'a' -Status 'active' -Roles @($exchangeAdmin)),
                    (New-Relationship -Id 'b' -Status 'terminated'),
                    (New-Relationship -Id 'c' -Status 'approvalPending' -Roles @($globalAdmin) -Created ([datetime]::UtcNow.AddDays(-80))),
                    (New-Relationship -Id 'd' -Status 'active' -Tenant '99999999-9999-9999-9999-999999999999' -Roles @('e8611ab8-c189-46e8-94e1-60213ab1f814'))
                )
            }
        )
    }
    It 'leaves out terminated relationships by default and reads the partner tenant' {
        $r = @(Get-MspGdapRelationship)
        $r.Id | Should -Not -Contain 'b'
        (Get-FakeGraphCall)[0].PartnerTenant | Should -BeTrue
    }
    It 'filters by status and customer' {
        @(Get-MspGdapRelationship -Status active -TenantId $customer).Id | Should -Be @('a')
    }
    It 'flags Global Administrator, privileged roles and ageing requests' {
        $c = Get-MspGdapRelationship | Where-Object Id -eq 'c'
        $c.ContainsGlobalAdministrator | Should -BeTrue
        $c.ContainsPrivilegedRoles | Should -BeTrue
        $c.ApprovalExpiresSoon | Should -BeTrue
        (Get-MspGdapRelationship | Where-Object Id -eq 'd').ContainsPrivilegedRoles | Should -BeTrue
        (Get-MspGdapRelationship | Where-Object Id -eq 'a').RoleNames | Should -Be @('Exchange Administrator')
    }
}

Describe 'New-MspGdapRelationship' {
    BeforeEach {
        $global:MspTest.State = @{ Created = $null; LockStatus = 'approvalPending' }
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern 'delegatedAdminRelationships$' -Response { @((New-Relationship -Id 'old' -Name 'Taken-Name' -Status 'terminated')) }
            New-FakeRoute -Method POST -Pattern 'delegatedAdminRelationships$' -Response {
                param($Body)
                $global:MspTest.State.Created = $Body
                [pscustomobject]@{ id = 'new-rel'; status = 'created' }
            }
            New-FakeRoute -Method POST -Pattern 'new-rel/requests$' -Response { [pscustomobject]@{ id = 'req-1'; status = 'created'; action = 'lockForApproval' } }
            New-FakeRoute -Pattern 'delegatedAdminRelationships/new-rel$' -Response {
                $body = $global:MspTest.State.Created
                $status = if ((Get-FakeGraphCall -Method POST -Pattern 'requests$').Count -gt 0) { $global:MspTest.State.LockStatus } else { 'created' }
                [pscustomobject]@{ id = 'new-rel'; status = $status; accessDetails = [pscustomobject]@{ unifiedRoles = @($body.accessDetails.unifiedRoles | ForEach-Object { [pscustomobject]@{ roleDefinitionId = $_.roleDefinitionId } }) } }
            }
        )
    }
    It 'previews with -WhatIf without writing, listing the default least-privilege roles' {
        $r = New-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com' -WhatIf
        $r.Outcome | Should -Be 'WhatIf'
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
        $r.Roles | Should -Contain 'Cloud Application Administrator'
        $r.Roles | Should -Not -Contain 'Global Administrator'
        $r.Roles | Should -Not -Contain 'Privileged Role Administrator'
    }
    It 'refuses Global Administrator without -IncludeGlobalAdministrator' {
        $r = New-MspGdapRelationship -TenantId $customer -Role 'Global Administrator', 'Exchange Administrator' -Confirm:$false
        $r.Success | Should -BeFalse
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'forces auto-extend off when Global Administrator is included' {
        $r = New-MspGdapRelationship -TenantId $customer -Role 'Global Administrator' -IncludeGlobalAdministrator -AutoExtend -Confirm:$false
        $global:MspTest.State.Created.autoExtendDuration | Should -Be 'PT0S'
        ($r.Steps | Where-Object Step -eq 'Auto-extend').Status | Should -Be 'Warning'
    }
    It 'rejects a duration over two years' {
        $r = New-MspGdapRelationship -TenantId $customer -Duration 'P3Y' -Confirm:$false
        ($r.Steps | Where-Object Step -eq 'Check duration').Status | Should -Be 'Failed'
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'accepts P730D and rejects P731D, the Graph maximum being two years' {
        (New-MspGdapRelationship -TenantId $customer -Duration 'P730D' -WhatIf).Steps | Where-Object Step -eq 'Check duration' | ForEach-Object Status | Should -Be 'Passed'
        $r = New-MspGdapRelationship -TenantId $customer -Duration 'P731D' -Confirm:$false -ErrorAction SilentlyContinue
        ($r.Steps | Where-Object Step -eq 'Check duration').Status | Should -Be 'Failed'
    }
    It 'refuses the partner tenant' {
        $r = New-MspGdapRelationship -TenantId $global:MspTest.PartnerTenantId -WhatIf
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'partner tenant'
        (Get-FakeGraphCall).Count | Should -Be 0
    }
    It 'refuses a role template ID that is not in the catalogue unless -AllowUnknownRole' {
        $r = New-MspGdapRelationship -TenantId $customer -Role 'abababab-abab-abab-abab-abababababab' -Confirm:$false -ErrorAction SilentlyContinue
        ($r.Steps | Where-Object Step -eq 'Resolve roles').Status | Should -Be 'Failed'
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
        $r2 = New-MspGdapRelationship -TenantId $customer -Role 'abababab-abab-abab-abab-abababababab' -AllowUnknownRole -WhatIf
        ($r2.Steps | Where-Object Step -eq 'Roles not in catalogue').Status | Should -Be 'Warning'
    }
    It 'warns about highly privileged roles' {
        $r = New-MspGdapRelationship -TenantId $customer -Role 'Privileged Role Administrator' -WhatIf
        ($r.Steps | Where-Object Step -eq 'Highly privileged roles').Status | Should -Be 'Warning'
    }
    It 'raises a non-terminating error when creation fails' {
        { New-MspGdapRelationship -TenantId $customer -DisplayName 'Taken-Name' -Confirm:$false -ErrorAction Stop } | Should -Throw -ErrorId 'MspGdap.New-MspGdapRelationship.Failed*'
    }
    It 'rejects a display name that is already used' {
        $r = New-MspGdapRelationship -TenantId $customer -DisplayName 'Taken-Name' -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Check display name').Detail | Should -Not -Match '365'
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'creates from the access map, locks for approval and confirms by readback' {
        $r = New-MspGdapRelationship -TenantId $customer -AccessMapPath $mapPath -AutoExtend -Confirm:$false
        $r.Success | Should -BeTrue
        $r.Status | Should -Be 'approvalPending'
        $body = $global:MspTest.State.Created
        $body.autoExtendDuration | Should -Be 'P180D'
        $body.customer.tenantId | Should -Be $customer
        @($body.accessDetails.unifiedRoles.roleDefinitionId) | Sort-Object | Should -Be (@($exchangeAdmin, $helpdesk, $readers) | Sort-Object)
        $body.displayName.Length | Should -BeLessOrEqual 50
        (Get-FakeGraphCall -Method POST -Pattern 'requests$')[0].Body.action | Should -Be 'lockForApproval'
    }
    It 'does not report success when the lock is never confirmed' {
        $global:MspTest.State.LockStatus = 'created'
        $r = New-MspGdapRelationship -TenantId $customer -Role 'Exchange Administrator' -Confirm:$false -LockTimeoutSeconds 10
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Lock for approval').Status | Should -Be 'Unknown'
    }
}

Describe 'Set-MspGdapAccessAssignment' {
    BeforeEach {
        $global:MspTest.State = @{ Relationship = (New-Relationship -Id 'rel-1'); Assignments = @{}; Posted = @(); ReadStatus = 'active' }
        Mock Invoke-MspGraphDirect { [pscustomobject]@{ StatusCode = 202; Body = $null; Error = $null } }
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern 'delegatedAdminRelationships/rel-1$' -Response { $global:MspTest.State.Relationship }
            New-FakeRoute -Pattern 'rel-1/accessAssignments$' -Response { @($global:MspTest.State.Assignments.Values) }
            New-FakeRoute -Method POST -Pattern 'rel-1/accessAssignments$' -Response {
                param($Body)
                $id = "as-$($Body.accessContainer.accessContainerId.Substring(0, 4))"
                $global:MspTest.State.Assignments[$id] = [pscustomobject]@{
                    id = $id; status = 'pending'; '@odata.etag' = 'W/"1"'
                    accessContainer = [pscustomobject]@{ accessContainerId = $Body.accessContainer.accessContainerId; accessContainerType = 'securityGroup' }
                    accessDetails = [pscustomobject]@{ unifiedRoles = @($Body.accessDetails.unifiedRoles | ForEach-Object { [pscustomobject]@{ roleDefinitionId = $_.roleDefinitionId } }) }
                }
                [pscustomobject]@{ id = $id; status = 'pending' }
            }
            New-FakeRoute -Pattern 'rel-1/accessAssignments/(?<id>[^/?]+)$' -Response {
                param($Body, $Uri)
                $id = ($Uri -split '/')[-1]
                $a = $global:MspTest.State.Assignments[$id]
                $a.status = $global:MspTest.State.ReadStatus
                $a
            }
        )
    }
    It 'refuses a relationship that is not active' {
        $global:MspTest.State.Relationship = New-Relationship -Id 'rel-1' -Status 'approvalPending'
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -Confirm:$false)
        $r[0].Success | Should -BeFalse
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'creates one assignment per group and confirms it is active' {
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -Confirm:$false)
        $r.Count | Should -Be 2
        $r.Success | Should -Not -Contain $false
        (Get-FakeGraphCall -Method POST).Count | Should -Be 2
        (Get-FakeGraphCall -Method POST)[0].Body.accessContainer.accessContainerType | Should -Be 'securityGroup'
    }
    It 'reports Unknown, not success, while an assignment stays pending' {
        $global:MspTest.State.ReadStatus = 'pending'
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -Confirm:$false -WaitSeconds 20)
        $r.Success | Should -Not -Contain $true
        ($r[0].Steps | Where-Object Step -eq 'Assign roles').Status | Should -Be 'Unknown'
    }
    It 'fails a group whose roles are not in the relationship, without writing' {
        $global:MspTest.State.Relationship = New-Relationship -Id 'rel-1' -Roles @($helpdesk, $readers)
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -Confirm:$false)
        ($r | Where-Object GroupId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb').Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 1
    }
    It 'leaves matching assignments alone and patches changed ones with If-Match' {
        $global:MspTest.State.Assignments['as-aaaa'] = [pscustomobject]@{ id = 'as-aaaa'; status = 'active'; '@odata.etag' = 'W/"7"'
            accessContainer = [pscustomobject]@{ accessContainerId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }
            accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = $helpdesk }, [pscustomobject]@{ roleDefinitionId = $readers }) } }
        $global:MspTest.State.Assignments['as-bbbb'] = [pscustomobject]@{ id = 'as-bbbb'; status = 'active'; '@odata.etag' = 'W/"9"'
            accessContainer = [pscustomobject]@{ accessContainerId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' }
            accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = $readers }) } }
        Mock Invoke-MspGraphDirect {
            $global:MspTest.State.Assignments['as-bbbb'].accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = $exchangeAdmin }) }
            [pscustomobject]@{ StatusCode = 202; Body = $null; Error = $null }
        }
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -Confirm:$false)
        ($r | Where-Object GroupId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa').Steps[0].Status | Should -Be 'Passed'
        ($r | Where-Object GroupId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb').Steps[-1].Status | Should -Be 'Changed'
        Should -Invoke Invoke-MspGraphDirect -Times 1 -ParameterFilter { $Method -eq 'PATCH' -and $Headers['If-Match'] -eq 'W/"9"' }
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'writes nothing with -WhatIf' {
        $r = @(Set-MspGdapAccessAssignment -RelationshipId 'rel-1' -AccessMapPath $mapPath -WhatIf)
        $r.Outcome | Should -Not -Contain 'Succeeded'
        (Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
}

Describe 'Test-MspGdapAccess' {
    BeforeEach {
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern 'delegatedAdminRelationships$' -Response { @((New-Relationship -Id 'rel-1')) }
            New-FakeRoute -Pattern 'rel-1/accessAssignments$' -Response {
                @(
                    [pscustomobject]@{ id = 'x'; status = 'active'; accessContainer = [pscustomobject]@{ accessContainerId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }; accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = $helpdesk }) } },
                    [pscustomobject]@{ id = 'y'; status = 'active'; accessContainer = [pscustomobject]@{ accessContainerId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' }; accessDetails = [pscustomobject]@{ unifiedRoles = @([pscustomobject]@{ roleDefinitionId = $exchangeAdmin }) } }
                )
            }
            New-FakeRoute -Pattern 'me/transitiveMemberOf' -Response { @([pscustomobject]@{ id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; displayName = 'GDAP-Helpdesk' }) }
            New-FakeRoute -Pattern 'organization' -Response { [pscustomobject]@{ id = $customer; displayName = 'Fabrikam' } }
        )
    }
    It 'proves the customer token but never asks the customer for /me/memberOf' {
        $r = Test-MspGdapAccess -TenantId $customer
        $r.Success | Should -BeTrue
        ($r.Steps | Where-Object Step -eq 'Customer token').Status | Should -Be 'Passed'
        $view = $r.Steps | Where-Object Step -eq 'Customer-side role view'
        $view.Status | Should -Be 'Skipped'
        $view.Detail | Should -Match 'not objects in the customer directory'
        $view.Detail | Should -Match 'Effective roles'
        @(Get-FakeGraphCall -Pattern 'memberOf' | Where-Object Uri -notmatch 'transitiveMemberOf').Count | Should -Be 0
        @(Get-FakeGraphCall | Where-Object { $_.TenantId -eq $customer }).Uri | Should -Not -Match 'me/'
    }
    It 'makes no customer call with -SkipCustomerProbe' {
        $r = Test-MspGdapAccess -TenantId $customer -SkipCustomerProbe
        $r.Success | Should -BeTrue
        @($r.Steps | Where-Object { $_.Step -in 'Customer token', 'Customer-side role view' }).Count | Should -Be 0
        @(Get-FakeGraphCall | Where-Object { $_.TenantId -eq $customer }).Count | Should -Be 0
    }
    It 'lists only roles from groups the technician is in' {
        $r = Test-MspGdapAccess -TenantId 'fabrikam.onmicrosoft.com'
        $r.EffectiveRoleNames | Should -Be @('Helpdesk Administrator')
        $r.Success | Should -BeTrue
    }
    It 'fails when a required role is missing' {
        $r = Test-MspGdapAccess -TenantId $customer -RequiredRole 'Exchange Administrator'
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Required roles').Detail | Should -Match 'Exchange Administrator'
    }
    It 'passes -AnyRole when one is held' {
        (Test-MspGdapAccess -TenantId $customer -AnyRole 'Cloud Application Administrator', 'Helpdesk Administrator').Success | Should -BeTrue
    }
    It 'fails when there is no active relationship' {
        Set-FakeGraphRoute @(New-FakeRoute -Pattern 'delegatedAdminRelationships$' -Response { @() })
        (Test-MspGdapAccess -TenantId $customer).Success | Should -BeFalse
    }
}
