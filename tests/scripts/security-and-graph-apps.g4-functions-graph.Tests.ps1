#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Mocked smoke tests for the scripts in scripts/security-and-graph-apps that replace the retired AzureAD,
# AzureRM, AdminAgents and Azure AD Graph articles (see scripts/MAPPING.json). Nothing here
# contacts a tenant or IT Glue: MspGdap commands, Microsoft Graph and Exchange Online cmdlets are mocked.

BeforeAll {
    $repoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath '..')).ProviderPath
    Import-Module (Join-Path -Path $repoRoot -ChildPath 'src' -AdditionalChildPath 'MspGdap', 'MspGdap.psd1') -Force
    $script:ScriptDir = Join-Path -Path $repoRoot -ChildPath 'scripts' -AdditionalChildPath 'security-and-graph-apps'
    $script:Tenant = 'contoso.onmicrosoft.com'
    $global:G4SecGraphCalls = [System.Collections.Generic.List[object]]::new()

    $stubNames = @(
        'Disconnect-ExchangeOnline', 'New-ComplianceSearch', 'Start-ComplianceSearch', 'Get-ComplianceSearch',
        'New-ComplianceSearchAction', 'Get-ComplianceSearchAction', 'Get-AdminAuditLogConfig', 'Search-UnifiedAuditLog', 'Get-Secret',
        'Push-OutputBinding'
    )
    # Leave a real Get-Secret (SecretManagement installed) in place. Shadowing it with a global stub and removing
    # the stub later confuses Pester's mock call history for test files that run afterwards and mock Get-Secret.
    $stubNames = @($stubNames | Where-Object { $_ -ne 'Get-Secret' -or -not (Get-Command -Name 'Get-Secret' -CommandType Cmdlet -ErrorAction SilentlyContinue) })
    foreach ($name in $stubNames) {
        $null = New-Item -Path "Function:\global:$name" -Value { [CmdletBinding()] param([Parameter(ValueFromRemainingArguments)][object[]]$Rest) } -Force
    }
    $script:StubNames = $stubNames

    function global:Get-G4SecurityGraphFixture {
        param([string]$Method, [string]$Uri)
        if ($Method -ne 'GET') { return [pscustomobject]@{ id = 'new-object-id'; appId = '33333333-3333-3333-3333-333333333333' } }
        $bom = [string][char]0xFEFF
        switch -Regex ($Uri) {
            '^organization' { return [pscustomobject]@{ id = '11111111-1111-1111-1111-111111111111'; displayName = 'Contoso'; verifiedDomains = @([pscustomobject]@{ name = 'contoso.com'; isDefault = $true; isInitial = $false }) } }
            '^security/secureScores' {
                return [pscustomobject]@{
                    createdDateTime          = '2026-10-06T00:00:00Z'
                    currentScore             = 40
                    maxScore                 = 80
                    licensedUserCount        = 10
                    activeUserCount          = 9
                    enabledServices          = @('HasExchange', 'HasSharePoint')
                    averageComparativeScores = @([pscustomobject]@{ basis = 'TotalSeats'; averageScore = 35.5 })
                    controlScores            = @(
                        [pscustomobject]@{ controlName = 'MFARegistrationV2'; controlCategory = 'Identity'; score = 5 }
                        [pscustomobject]@{ controlName = 'DLPEnabled'; controlCategory = 'Data'; score = 0 }
                    )
                }
            }
            '^security/secureScoreControlProfiles' {
                return @(
                    [pscustomobject]@{ id = 'MFARegistrationV2'; title = 'Ensure all users can complete MFA'; maxScore = 9; rank = 1; userImpact = 'Moderate'; implementationCost = 'Low'; actionUrl = 'https://security.microsoft.com' }
                    [pscustomobject]@{ id = 'DLPEnabled'; title = 'Turn on DLP policies'; maxScore = 1; rank = 2; userImpact = 'Low'; implementationCost = 'Low'; actionUrl = 'https://purview.microsoft.com' }
                )
            }
            '^admin/serviceAnnouncement/messages' {
                return @(
                    [pscustomobject]@{ id = 'MC1'; title = 'Retirement of a legacy feature'; category = 'planForChange'; severity = 'normal'; services = @('Exchange Online'); isMajorChange = $true; actionRequiredByDateTime = '2026-12-01T00:00:00Z'; lastModifiedDateTime = [datetime]::UtcNow.AddDays(-2); body = [pscustomobject]@{ content = 'Details' } }
                    [pscustomobject]@{ id = 'MC2'; title = 'New feature'; category = 'stayInformed'; severity = 'normal'; services = @('Teams'); isMajorChange = $false; actionRequiredByDateTime = $null; lastModifiedDateTime = [datetime]::UtcNow.AddDays(-40); body = [pscustomobject]@{ content = 'Nothing to do' } }
                )
            }
            '^identityProtection/riskDetections' {
                return @(
                    [pscustomobject]@{ id = 'd1'; userPrincipalName = 'jane@contoso.com'; userDisplayName = 'Jane Citizen'; riskEventType = 'unfamiliarFeatures'; riskLevel = 'medium'; riskState = 'atRisk'; detectedDateTime = '2026-10-06T00:00:00Z'; ipAddress = '203.0.113.5'; location = [pscustomobject]@{ city = 'Sydney'; state = 'NSW'; countryOrRegion = 'AU' } }
                    [pscustomobject]@{ id = 'd2'; userPrincipalName = 'joe@contoso.com'; riskEventType = 'anonymizedIPAddress'; riskLevel = 'low'; riskState = 'remediated'; detectedDateTime = '2026-10-06T00:00:00Z' }
                )
            }
            '^identityProtection/riskyUsers' { return [pscustomobject]@{ id = 'u1'; userPrincipalName = 'jane@contoso.com'; riskLevel = 'high'; riskState = 'atRisk'; riskLastUpdatedDateTime = '2026-10-06T00:00:00Z' } }
            'getMailboxUsageDetail' {
                return $bom + "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Item Count,Storage Used (Byte),Prohibit Send/Receive Quota (Byte),Recipient Type,Has Archive,Last Activity Date`r`n2026-10-06,jane@contoso.com,Jane,False,100,48318382080,53687091200,User,False,2026-10-05`r`n2026-10-06,joe@contoso.com,Joe,False,10,1073741824,53687091200,User,False,2026-10-05`r`n"
            }
            'getOffice365ActiveUserDetail' {
                return "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Assigned Products,Exchange Last Activity Date,OneDrive Last Activity Date,SharePoint Last Activity Date,Teams Last Activity Date`n2026-10-06,jane@contoso.com,Jane,False,MICROSOFT 365 BUSINESS PREMIUM,2026-10-05,2026-10-04,,2026-10-05`n"
            }
            'getEmailActivityUserDetail' {
                return "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Send Count,Receive Count,Read Count`n2026-10-06,jane@contoso.com,Jane,False,12,340,300`n"
            }
            'getTeamsUserActivityUserDetail' {
                return "Report Refresh Date,User Principal Name,Is Deleted,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count`n2026-10-06,jane@contoso.com,False,5,20,3,4`n"
            }
            'getOneDriveUsageAccountDetail' {
                return "Report Refresh Date,Owner Principal Name,Storage Used (Byte),File Count`n2026-10-06,jane@contoso.com,2147483648,120`n"
            }
            '^security/alerts_v2/' { return [pscustomobject]@{ id = 'alert1'; title = 'Suspicious sign-in'; severity = 'medium'; status = 'new'; serviceSource = 'microsoftDefenderForOffice365' } }
            '^security/alerts_v2\?' {
                return @(
                    [pscustomobject]@{ id = 'alert1'; title = 'Suspicious sign-in'; severity = 'medium'; status = 'new'; createdDateTime = '2026-10-06T00:00:00Z' }
                    [pscustomobject]@{ id = 'alert2'; title = 'Old'; severity = 'high'; status = 'resolved'; createdDateTime = '2026-10-05T00:00:00Z' }
                    [pscustomobject]@{ id = 'alert3'; title = 'Info'; severity = 'informational'; status = 'new'; createdDateTime = '2026-10-05T00:00:00Z' }
                )
            }
            '^sites/' { return [pscustomobject]@{ id = 'contoso.sharepoint.com,site-guid,web-guid' } }
            '^applications\?' { return @() }
            '^servicePrincipals\(appId=' { return [pscustomobject]@{ id = 'graph-sp-id'; appRoles = @([pscustomobject]@{ id = 'sites-selected-role-id'; value = 'Sites.Selected'; allowedMemberTypes = @('Application') }) } }
            '^auditLogs/signIns' { return [pscustomobject]@{ ipAddress = '203.0.113.5'; location = [pscustomobject]@{ city = 'Sydney'; state = 'NSW'; countryOrRegion = 'AU' } } }
            default { return $null }
        }
    }

    function Assert-G4CustomerRow {
        param([object[]]$Rows, [string[]]$Property)
        $Rows.Count | Should -BeGreaterThan 0
        foreach ($row in $Rows) {
            $row.CustomerTenantId | Should -Be '11111111-1111-1111-1111-111111111111'
            $row.CustomerName | Should -Be 'Contoso'
            $row.Error | Should -BeNullOrEmpty
            foreach ($name in $Property) { $row.PSObject.Properties.Name | Should -Contain $name }
        }
    }
}

AfterAll {
    foreach ($name in $script:StubNames) { Remove-Item -Path "Function:\global:$name" -ErrorAction SilentlyContinue }
    Remove-Item -Path 'Function:\global:Get-G4SecurityGraphFixture' -ErrorAction SilentlyContinue
    Remove-Variable -Name G4SecGraphCalls, G4SecUalServed, G4SecBindings, G4SecSpAttempts -Scope Global -ErrorAction SilentlyContinue
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'scripts/security-and-graph-apps (g4 Graph app rewrites)' {
    BeforeEach {
        $global:G4SecGraphCalls.Clear()
        Mock Invoke-MspGraphRequest {
            $verb = if ($Method) { $Method } else { 'GET' }
            $global:G4SecGraphCalls.Add([pscustomobject]@{ Method = $verb; Uri = $Uri; TenantId = $TenantId; Body = $Body })
            Get-G4SecurityGraphFixture -Method $verb -Uri $Uri
        }
        Mock Get-MspCustomer { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; DisplayName = 'Contoso'; DefaultDomainName = 'contoso.com' } }
        Mock Connect-MspExchangeOnline { [pscustomobject]@{ TenantId = $TenantId; Mode = 'Delegated' } }
        Mock Connect-MspSecurityCompliance { [pscustomobject]@{ TenantId = $TenantId } }
        Mock Disconnect-ExchangeOnline {}
        Mock Start-Sleep {}
        Mock New-ComplianceSearch { [pscustomobject]@{ Name = 'search' } }
        Mock Start-ComplianceSearch {}
        Mock Get-ComplianceSearch { [pscustomobject]@{ Status = 'Completed'; Items = 3; Size = 30000; SuccessResults = "{Location: jane@contoso.com, Item count: 2, Total size: 20000`r`nLocation: joe@contoso.com, Item count: 1, Total size: 10000`r`nLocation: amy@contoso.com, Item count: 0, Total size: 0}" } }
        Mock New-ComplianceSearchAction {}
        Mock Get-ComplianceSearchAction { [pscustomobject]@{ Status = 'Completed' } }
        Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
        Mock Search-UnifiedAuditLog {
            if ($global:G4SecUalServed) { return @() }
            $global:G4SecUalServed = $true
            @(
                [pscustomobject]@{ Identity = 'r1'; ResultCount = 2; CreationDate = '2026-10-06'; Operations = 'MailItemsAccessed'; AuditData = '{"CreationTime":"2026-10-06T01:00:00","Workload":"Exchange","Operation":"MailItemsAccessed","ResultStatus":"Succeeded","ClientIP":"[203.0.113.5]:443"}' }
                [pscustomobject]@{ Identity = 'r2'; ResultCount = 2; CreationDate = '2026-10-06'; Operations = 'New-InboxRule'; AuditData = '{"CreationTime":"2026-10-06T02:00:00","Workload":"Exchange","Operation":"New-InboxRule","ResultStatus":"True","ClientIP":"198.51.100.7:51000"}' }
            )
        }
        Mock Test-MspPartnerAppConsent { [pscustomobject]@{ Success = $false; Outcome = 'Failed'; Steps = @([pscustomobject]@{ Step = 'Graph scopes'; Status = 'Failed'; Detail = 'User.Read.All missing' }) } }
        Mock Grant-MspPartnerAppConsent { [pscustomobject]@{ Success = $true; Outcome = 'Succeeded'; Steps = @() } }
        $global:G4SecUalServed = $false
    }

    Context 'get-secure-scores.ps1' {
        It 'returns one score row per customer' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-secure-scores.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'CurrentScore', 'MaxScore', 'Percentage', 'SimilarTenantsAverage'
            $rows[0].Percentage | Should -Be 50
            $rows[0].SimilarTenantsAverage | Should -Be 35.5
            $rows[0].IdentityScore | Should -Be 5
            $rows[0].DataScore | Should -Be 0
        }

        It 'returns one row per control with -IncludeControls' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-secure-scores.ps1') -AllCustomers -IncludeControls)
            Assert-G4CustomerRow -Rows $rows -Property 'ControlName', 'Title', 'Score', 'MaxScore'
            $rows.Count | Should -Be 2
            ($rows | Where-Object ControlName -eq 'MFARegistrationV2').Title | Should -Be 'Ensure all users can complete MFA'
        }

        It 'records a failed customer and carries on' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $TenantId -eq 'broken.onmicrosoft.com' } -MockWith { throw 'Forbidden' }
            $rows = @(& (Join-Path $script:ScriptDir 'get-secure-scores.ps1') -TenantId 'broken.onmicrosoft.com', $script:Tenant -WarningAction SilentlyContinue)
            $rows.Count | Should -Be 2
            ($rows | Where-Object CustomerTenantId -eq 'broken.onmicrosoft.com').Error | Should -Be 'Forbidden'
        }
    }

    Context 'get-message-center-posts.ps1' {
        It 'filters by text, action date and age' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-message-center-posts.ps1') -TenantId $script:Tenant -SearchText 'retirement' -ActionRequiredOnly -Days 30)
            Assert-G4CustomerRow -Rows $rows -Property 'MessageId', 'Title', 'ActionRequiredByDateTime', 'Services'
            $rows.Count | Should -Be 1
            $rows[0].MessageId | Should -Be 'MC1'
            $rows[0].Body | Should -Be 'Details'
        }
    }

    Context 'get-risk-detections.ps1' {
        It 'returns detections at or above the minimum level and risky users' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-risk-detections.ps1') -TenantId $script:Tenant -MinimumRiskLevel medium -IncludeRiskyUsers)
            Assert-G4CustomerRow -Rows $rows -Property 'RecordType', 'UserPrincipalName', 'RiskLevel', 'Location'
            $rows.Count | Should -Be 2
            ($rows | Where-Object RecordType -eq 'RiskDetection').Location | Should -Be 'Sydney, NSW, AU'
            ($rows | Where-Object RecordType -eq 'RiskDetection').UserDisplayName | Should -Be 'Jane Citizen'
            ($rows | Where-Object RecordType -eq 'RiskyUser').RiskLevel | Should -Be 'high'
            ($global:G4SecGraphCalls | Where-Object Uri -like 'identityProtection/riskDetections*').Uri | Should -Match 'detectedDateTime ge '
        }
    }

    Context 'get-mailbox-usage.ps1' {
        It 'parses the CSV report and flags nearly full mailboxes' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-mailbox-usage.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'UserPrincipalName', 'StorageUsedGB', 'PercentUsed', 'OverThreshold', 'NamesConcealed'
            $rows.Count | Should -Be 2
            $jane = $rows | Where-Object UserPrincipalName -eq 'jane@contoso.com'
            $jane.PercentUsed | Should -Be 90
            $jane.OverThreshold | Should -BeTrue
            $jane.NamesConcealed | Should -BeFalse
        }

        It 'returns only flagged mailboxes with -OverThresholdOnly' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-mailbox-usage.ps1') -AllCustomers -OverThresholdOnly)
            $rows.Count | Should -Be 1
        }

        It 'leaves out excluded mailboxes' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-mailbox-usage.ps1') -TenantId $script:Tenant -OverThresholdOnly -ExcludeUserPrincipalName 'jane@contoso.com')
            $rows.Count | Should -Be 0
        }
    }

    Context 'get-security-alerts.ps1' {
        It 'lists open alerts at or above the minimum severity' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-security-alerts.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'AlertId', 'Title', 'Severity', 'Status', 'AlertWebUrl'
            $rows.Count | Should -Be 1
            $rows[0].AlertId | Should -Be 'alert1'
        }

        It 'does not resolve alerts under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-security-alerts.ps1') -TenantId $script:Tenant -ResolveAlertId 'alert1' -Comment 'Expected' -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            @($global:G4SecGraphCalls | Where-Object Method -ne 'GET').Count | Should -Be 0
        }

        It 'resolves and comments with -Apply' {
            $null = & (Join-Path $script:ScriptDir 'get-security-alerts.ps1') -TenantId $script:Tenant -ResolveAlertId 'alert1' -Classification falsePositive -Comment 'Expected' -Apply -Confirm:$false
            @($global:G4SecGraphCalls | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -eq 'security/alerts_v2/alert1' }).Count | Should -Be 1
            @($global:G4SecGraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq 'security/alerts_v2/alert1/comments' }).Count | Should -Be 1
        }

        It 'queues each alert once from the SecurityAlertSync function' {
            $app = Join-Path $TestDrive 'app'
            $null = New-Item -ItemType Directory -Path (Join-Path $app 'scripts'), (Join-Path $app 'SecurityAlertSync') -Force
            Copy-Item -Path (Join-Path $script:ScriptDir 'get-security-alerts.ps1') -Destination (Join-Path $app 'scripts')
            Copy-Item -Path (Join-Path $script:ScriptDir 'functions' -AdditionalChildPath 'SecurityAlertSync', 'run.ps1') -Destination (Join-Path $app 'SecurityAlertSync')
            $global:G4SecBindings = @{}
            Mock Push-OutputBinding { $global:G4SecBindings[$Rest[1]] = $Rest[3] }
            $timer = [pscustomobject]@{ IsPastDue = $false }
            $env:MSPGDAP_TENANT_IDS = $script:Tenant
            try {
                & (Join-Path $app 'SecurityAlertSync' -AdditionalChildPath 'run.ps1') -Timer $timer -previousState $null
                @($global:G4SecBindings['alerts']).Count | Should -Be 1
                $state = $global:G4SecBindings['state']
                $state | Should -Match 'alert1'

                $global:G4SecBindings = @{}
                & (Join-Path $app 'SecurityAlertSync' -AdditionalChildPath 'run.ps1') -Timer $timer -previousState $state
                $global:G4SecBindings.ContainsKey('alerts') | Should -BeFalse
            }
            finally {
                Remove-Item -Path Env:\MSPGDAP_TENANT_IDS -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'new-sharepoint-site-app.ps1' {
        BeforeAll {
            $rsa = [System.Security.Cryptography.RSA]::Create(2048)
            $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=Contoso Reports Writer', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $certificate = $request.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(30))
            $script:CerPath = Join-Path $TestDrive 'reports-writer.cer'
            [System.IO.File]::WriteAllBytes($script:CerPath, $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
            $certificate.Dispose()
            $rsa.Dispose()
        }

        It 'creates nothing under -WhatIf' {
            $params = @{ TenantId = $script:Tenant; DisplayName = 'Contoso Reports Writer'; SiteUrl = 'https://contoso.sharepoint.com/sites/Reports'; Role = 'write'; CertificatePath = $script:CerPath; Apply = $true; WhatIf = $true }
            $rows = @(& (Join-Path $script:ScriptDir 'new-sharepoint-site-app.ps1') @params)
            Assert-G4CustomerRow -Rows $rows -Property 'SiteId', 'AppId', 'Role', 'CertificateThumbprint', 'Action'
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].SiteId | Should -Be 'contoso.sharepoint.com,site-guid,web-guid'
            @($global:G4SecGraphCalls | Where-Object Method -ne 'GET').Count | Should -Be 0
        }

        It 'creates the app with Sites.Selected, a certificate and one site permission with -Apply' {
            $params = @{ TenantId = $script:Tenant; DisplayName = 'Contoso Reports Writer'; SiteUrl = 'https://contoso.sharepoint.com/sites/Reports'; Role = 'write'; CertificatePath = $script:CerPath; Apply = $true; Confirm = $false }
            $rows = @(& (Join-Path $script:ScriptDir 'new-sharepoint-site-app.ps1') @params)
            $rows[0].Action | Should -Be 'Created'
            $app = $global:G4SecGraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq 'applications' }
            $app.Body.requiredResourceAccess[0].resourceAccess[0].id | Should -Be 'sites-selected-role-id'
            $app.Body.keyCredentials[0].type | Should -Be 'AsymmetricX509Cert'
            $app.Body.Keys | Should -Not -Contain 'passwordCredentials'
            $permission = $global:G4SecGraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -like 'sites/*/permissions' }
            $permission.Body.roles | Should -Be @('write')
        }

        It 'retries the service principal while the new app replicates' {
            $global:G4SecSpAttempts = 0
            Mock Invoke-MspGraphRequest -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'servicePrincipals' } -MockWith {
                $global:G4SecSpAttempts++
                if ($global:G4SecSpAttempts -lt 2) { throw 'Resource does not exist yet' }
                [pscustomobject]@{ id = 'new-sp-id' }
            }
            $params = @{ TenantId = $script:Tenant; DisplayName = 'Contoso Reports Writer'; SiteUrl = 'https://contoso.sharepoint.com/sites/Reports'; Role = 'read'; CertificatePath = $script:CerPath; Apply = $true; Confirm = $false }
            $rows = @(& (Join-Path $script:ScriptDir 'new-sharepoint-site-app.ps1') @params)
            $rows[0].Action | Should -Be 'Created'
            $rows[0].ServicePrincipalId | Should -Be 'new-sp-id'
            $global:G4SecSpAttempts | Should -Be 2
        }

        It 'returns the app ID when a later step fails' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $Method -eq 'POST' -and $Uri -like 'sites/*/permissions' } -MockWith { throw 'Forbidden' }
            $params = @{ TenantId = $script:Tenant; DisplayName = 'Contoso Reports Writer'; SiteUrl = 'https://contoso.sharepoint.com/sites/Reports'; Role = 'read'; CertificatePath = $script:CerPath; Apply = $true; Confirm = $false }
            $rows = @(& (Join-Path $script:ScriptDir 'new-sharepoint-site-app.ps1') @params -WarningAction SilentlyContinue)
            $rows[0].Action | Should -Be 'FailedAfterAppCreated'
            $rows[0].AppId | Should -Be '33333333-3333-3333-3333-333333333333'
        }
    }

    Context 'sync-partner-app-consent.ps1' {
        It 'reports missing consent and grants nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'sync-partner-app-consent.ps1') -AllCustomers -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'Consented', 'Outcome', 'Action', 'Detail'
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].Detail | Should -Match 'User.Read.All missing'
            Should -Invoke Grant-MspPartnerAppConsent -Times 0 -Exactly
        }

        It 'grants consent with -Apply' {
            $rows = @(& (Join-Path $script:ScriptDir 'sync-partner-app-consent.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            $rows[0].Action | Should -Be 'Granted'
            Should -Invoke Grant-MspPartnerAppConsent -Times 1 -Exactly
        }

        It 'reports a customer that has no consent yet instead of failing it' {
            Mock Invoke-MspGraphRequest { throw 'AADSTS65001: consent missing' }
            $rows = @(& (Join-Path $script:ScriptDir 'sync-partner-app-consent.ps1') -TenantId $script:Tenant)
            $rows[0].Action | Should -Be 'WouldGrant'
            $rows[0].CustomerName | Should -Be 'Contoso'
            $rows[0].Error | Should -BeNullOrEmpty
        }
    }

    Context 'sync-secure-score-itglue.ps1' {
        It 'summarises scores by category without IT Glue' {
            $rows = @(& (Join-Path $script:ScriptDir 'sync-secure-score-itglue.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'SecureScore', 'IdentityScore', 'DataScore', 'Action'
            $rows[0].IdentityScore | Should -Be 5
            $rows[0].Action | Should -Be 'Collected'
        }

        It 'does not write to IT Glue under -WhatIf' {
            $mapPath = Join-Path $TestDrive 'map.csv'
            "TenantId,OrganizationId`n11111111-1111-1111-1111-111111111111,42" | Set-Content -Path $mapPath
            Mock Get-Secret { 'not-a-real-key' }
            Mock Invoke-RestMethod { [pscustomobject]@{ data = @([pscustomobject]@{ id = '9'; attributes = [pscustomobject]@{ traits = [pscustomobject]@{ 'tenant-id' = '11111111-1111-1111-1111-111111111111' } } }); links = $null } }
            $rows = @(& (Join-Path $script:ScriptDir 'sync-secure-score-itglue.ps1') -TenantId $script:Tenant -ITGlueFlexibleAssetTypeId 7 -ITGlueOrganizationMapPath $mapPath -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -ne 'GET' } -Times 0 -Exactly
        }

        It 'matches the IT Glue organisation by contact email domain when there is no map' {
            Mock Get-Secret { 'not-a-real-key' }
            Mock Invoke-RestMethod -ParameterFilter { "$Uri" -like '*/contacts*' } -MockWith {
                [pscustomobject]@{ data = @([pscustomobject]@{ attributes = [pscustomobject]@{ 'organization-id' = 77; 'contact-emails' = @([pscustomobject]@{ value = 'jane@contoso.com' }) } }); links = $null }
            }
            Mock Invoke-RestMethod -ParameterFilter { "$Uri" -like '*/flexible_assets*' } -MockWith { [pscustomobject]@{ data = @(); links = $null } }
            $rows = @(& (Join-Path $script:ScriptDir 'sync-secure-score-itglue.ps1') -TenantId $script:Tenant -ITGlueFlexibleAssetTypeId 7 -Apply -WhatIf)
            $rows[0].ITGlueOrganizationId | Should -Be '77'
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -ne 'GET' } -Times 0 -Exactly
        }
    }

    Context 'sync-user-activity-itglue.ps1' {
        It 'joins the usage reports by user' {
            $rows = @(& (Join-Path $script:ScriptDir 'sync-user-activity-itglue.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'UserPrincipalName', 'TeamsLastActivity', 'MailboxStorageGB', 'OneDriveStorageGB', 'Action'
            $rows.Count | Should -Be 1
            $rows[0].MailboxStorageGB | Should -Be 45
            $rows[0].OneDriveStorageGB | Should -Be 2
            $rows[0].OneDriveFileCount | Should -Be '120'
            $rows[0].EmailsSent | Should -Be '12'
            $rows[0].TeamsChatMessages | Should -Be 25
            @($global:G4SecGraphCalls | Where-Object Uri -like "reports/*(period='D90')").Count | Should -Be 7
        }
    }

    Context 'remove-email-from-mailboxes.ps1' {
        It 'creates no search and purges nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'remove-email-from-mailboxes.ps1') -TenantId $script:Tenant -Subject 'Invoice overdue' -From 'billing@fabrikam.example' -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'SearchName', 'Query', 'Items', 'Action'
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].Query | Should -Be 'subject:"Invoice overdue" AND from:billing@fabrikam.example'
            Should -Invoke New-ComplianceSearch -Times 0 -Exactly
            Should -Invoke New-ComplianceSearchAction -Times 0 -Exactly
        }

        It 'searches and reports without purging when -Apply is not used' {
            $rows = @(& (Join-Path $script:ScriptDir 'remove-email-from-mailboxes.ps1') -TenantId $script:Tenant -Subject 'Invoice overdue')
            $rows[0].Action | Should -Be 'Found'
            $rows[0].Items | Should -Be 3
            $rows[0].MailboxesWithHits | Should -Be 'jane@contoso.com; joe@contoso.com'
            Should -Invoke New-ComplianceSearch -Times 1 -Exactly
            Should -Invoke New-ComplianceSearchAction -Times 0 -Exactly
        }

        It 'refuses a query with only a date range' {
            { & (Join-Path $script:ScriptDir 'remove-email-from-mailboxes.ps1') -TenantId $script:Tenant -ReceivedAfter (Get-Date).AddDays(-1) } | Should -Throw '*date range alone*'
        }
    }

    Context 'get-user-breach-indicators.ps1' {
        It 'returns audit records with normalised IPs and sign-in locations' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-user-breach-indicators.ps1') -TenantId $script:Tenant -UserPrincipalName 'jane@contoso.com' -IncludeSignInLocations)
            Assert-G4CustomerRow -Rows $rows -Property 'CreationTime', 'Operation', 'ClientIP', 'Location', 'AuditLogEnabled'
            $rows.Count | Should -Be 2
            ($rows | Where-Object Operation -eq 'MailItemsAccessed').ClientIP | Should -Be '203.0.113.5'
            ($rows | Where-Object Operation -eq 'MailItemsAccessed').Location | Should -Be 'Sydney, NSW, AU'
            ($rows | Where-Object Operation -eq 'New-InboxRule').ClientIP | Should -Be '198.51.100.7'
            ($rows | Where-Object Operation -eq 'New-InboxRule').AuditData | Should -Match '"Operation":"New-InboxRule"'
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }
    }
}
