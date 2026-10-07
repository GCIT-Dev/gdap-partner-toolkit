#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Static checks over every script in scripts/ (the rewritten KB article scripts), plus scripts/MAPPING.json
# and scripts/README.md. Nothing here runs a script or contacts a tenant. The mocked smoke runs of each
# script are in tests/scripts.
#
# Standalone scripts (scripts/<family>/<name>.ps1) must:
#   - parse, and require PowerShell 7.4 and the MspGdap module
#   - have comment-based help with a synopsis, a description, every parameter, two examples, notes that name
#     the retired method, the GDAP roles and the partner app permissions, and links to the original article
#     and to the toolkit docs
#   - expose -TenantId (string array, pipeline), unless they are listed as partner tenant only below
#   - use SupportsShouldProcess and $PSCmdlet.ShouldProcess when they change anything
# Every .ps1 under scripts/ (including the Azure Functions run.ps1 and profile.ps1 files) must parse and must
# not contain any retired or unsafe method outside comments.

BeforeDiscovery {
    $scriptsRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..' 'scripts')).Path
    $allFiles = @(Get-ChildItem -Path $scriptsRoot -Recurse -Filter '*.ps1' -File)
    $script:AllScripts = foreach ($file in $allFiles) {
        @{ Name = [System.IO.Path]::GetRelativePath($scriptsRoot, $file.FullName).Replace('\', '/'); Path = $file.FullName }
    }
    $script:StandaloneScripts = foreach ($file in $allFiles) {
        if ($file.Directory.Parent.FullName.TrimEnd('\', '/') -eq $scriptsRoot.TrimEnd('\', '/')) {
            @{ Name = [System.IO.Path]::GetRelativePath($scriptsRoot, $file.FullName).Replace('\', '/'); Path = $file.FullName }
        }
    }
}

BeforeAll {
    $script:ScriptsRoot = (Resolve-Path -Path (Join-Path $PSScriptRoot '..' 'scripts')).Path

    # Scripts that only work in the partner tenant (IT Glue, UniFi and SharePoint lists), so -TenantId would
    # have nothing to act on. Each must say so in its help.
    $script:PartnerTenantOnly = @(
        'automation/itglue-quick-notes.ps1'
        'automation/sync-itglue-organisations-sharepoint.ps1'
        'automation/sync-unifi-devices-itglue.ps1'
    )

    # Retired or unsafe methods. Checked against the code with comments removed, because each script's help
    # names the method it replaces.
    $script:ForbiddenPatterns = [ordered]@{
        'MSOnline cmdlets'                        = '\b\w+-Msol[A-Z]\w*'
        'MSOnline module'                         = '\bMSOnline\b'
        'AzureAD module cmdlets'                  = '\b\w+-AzureAD\w*'
        'AzureAD module-qualified calls'          = '\bAzureAD(Preview)?\\'
        'AzureRM'                                 = 'AzureRM'
        'Exchange remote PowerShell session'      = 'New-PSSession[^\r\n]*Microsoft\.Exchange|ConfigurationName\s+[''"]?Microsoft\.Exchange'
        'Exchange basic authentication endpoint'  = 'powershell-liveid|DelegatedOrg='
        'Basic authentication'                    = '-Authentication\s+[''"]?Basic'
        'plain-text SecureString'                 = 'ConvertTo-SecureString[^\r\n]*-AsPlainText'
        'AES key files'                           = '(ConvertFrom|ConvertTo)-SecureString[^\r\n]*-Key\b'
        'password grant (ROPC)'                   = 'grant_type[''"]?\s*[=:]\s*[''"]?password'
        'Azure AD Graph'                          = 'graph\.windows\.net'
        'community PartnerCenter token command'   = 'New-PartnerAccessToken'
    }

    # Commands that change state somewhere. Anything matching these verbs counts as a write unless it is
    # local housekeeping in the allow list.
    $script:WriteVerbPattern = '^(Set|New|Remove|Add|Enable|Disable|Grant|Revoke|Update|Clear|Start|Stop|Register|Unregister|Install|Restore|Move|Rename|Send|Publish)-'
    $script:LocalCommands = @(
        'New-Item', 'New-Object', 'New-TimeSpan', 'New-Guid', 'New-Variable', 'Set-Variable', 'Remove-Variable',
        'Clear-Variable', 'Set-StrictMode', 'Add-Member', 'Add-Type', 'Set-Content', 'Add-Content', 'Set-Location',
        'Start-Sleep', 'Remove-Module', 'Register-ArgumentCompleter', 'Clear-MspTokenCache'
    )

    function Get-ScriptAst {
        param([string]$Path)
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
        [pscustomobject]@{ Ast = $ast; Tokens = $tokens; Errors = $errors }
    }

    function Get-CodeWithoutComment {
        param([string]$Path)
        $parsed = Get-ScriptAst -Path $Path
        $text = [System.Text.StringBuilder]::new((Get-Content -LiteralPath $Path -Raw))
        foreach ($token in $parsed.Tokens) {
            if ($token.Kind -eq [System.Management.Automation.Language.TokenKind]::Comment) {
                $start = $token.Extent.StartOffset
                $length = $token.Extent.EndOffset - $start
                $null = $text.Remove($start, $length).Insert($start, (' ' * $length))
            }
        }
        $text.ToString()
    }

    function Get-WriteCall {
        param([System.Management.Automation.Language.Ast]$Ast)
        $writeMethods = @('POST', 'PATCH', 'PUT', 'DELETE', 'MERGE')
        $commands = $Ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($command in $commands) {
            $name = $command.GetCommandName()
            if (-not $name) { continue }
            if ($name -match $script:WriteVerbPattern -and $script:LocalCommands -notcontains $name) {
                $name
                continue
            }
            # REST and Graph calls with a constant write method
            $elements = $command.CommandElements
            for ($i = 0; $i -lt $elements.Count - 1; $i++) {
                $element = $elements[$i]
                if ($element -is [System.Management.Automation.Language.CommandParameterAst] -and $element.ParameterName -eq 'Method') {
                    $argument = if ($element.Argument) { $element.Argument } else { $elements[$i + 1] }
                    if ($argument -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $writeMethods -contains $argument.Value.ToUpperInvariant()) {
                        "$name -Method $($argument.Value)"
                    }
                }
            }
        }
        # Splatted or batched requests: hashtable entries such as @{ Method = 'PATCH' }
        $tables = $Ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.HashtableAst] }, $true)
        foreach ($table in $tables) {
            foreach ($pair in $table.KeyValuePairs) {
                $key = $pair.Item1.Extent.Text.Trim('''', '"')
                if ($key -ne 'Method') { continue }
                $value = $pair.Item2.Extent.Text.Trim().Trim('''', '"')
                if ($writeMethods -contains $value.ToUpperInvariant()) { "hashtable Method = $value" }
            }
        }
    }
}

Describe 'Every script file: <Name>' -ForEach $script:AllScripts {
    BeforeAll {
        $script:Parsed = Get-ScriptAst -Path $Path
        $script:Code = Get-CodeWithoutComment -Path $Path
        $script:Lines = @(Get-Content -LiteralPath $Path)
    }

    It 'parses without errors' {
        @($script:Parsed.Errors | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Message)" }) | Should -BeNullOrEmpty
    }

    It 'uses no retired or unsafe method (<_>)' -ForEach @(
        'MSOnline cmdlets', 'MSOnline module', 'AzureAD module cmdlets', 'AzureAD module-qualified calls', 'AzureRM',
        'Exchange remote PowerShell session', 'Exchange basic authentication endpoint', 'Basic authentication',
        'plain-text SecureString', 'AES key files', 'password grant (ROPC)', 'Azure AD Graph', 'community PartnerCenter token command'
    ) {
        $script:Code | Should -Not -Match $script:ForbiddenPatterns[$_]
    }

    It 'has no em or en dashes' {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Not -Match "[$([char]0x2013)$([char]0x2014)]"
    }

    It 'mentions gcit.com.au only as a .LINK to the original article' {
        $bad = for ($i = 0; $i -lt $script:Lines.Count; $i++) {
            if ($script:Lines[$i] -notmatch 'gcit\.com\.au') { continue }
            $previous = $i - 1
            while ($previous -ge 0 -and -not $script:Lines[$previous].Trim()) { $previous-- }
            $isLink = $previous -ge 0 -and $script:Lines[$previous].Trim() -eq '.LINK' -and $script:Lines[$i].Trim() -match '^https://gcit\.com\.au/\S+$'
            if (-not $isLink) { "line $($i + 1)" }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Standalone script: <Name>' -ForEach $script:StandaloneScripts {
    BeforeAll {
        $script:Parsed = Get-ScriptAst -Path $Path
        $script:Ast = $script:Parsed.Ast
        $script:Help = $script:Ast.GetHelpContent()
        $script:Content = Get-Content -LiteralPath $Path -Raw
        $script:ParamNames = @($script:Ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $script:Binding = $script:Ast.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' } | Select-Object -First 1
    }

    It 'requires PowerShell 7.4 or later' {
        $script:Ast.ScriptRequirements.RequiredPSVersion | Should -BeGreaterOrEqual ([version]'7.4')
    }

    It 'requires the MspGdap module' {
        @($script:Ast.ScriptRequirements.RequiredModules | ForEach-Object Name) | Should -Contain 'MspGdap'
    }

    It 'is an advanced script with a param block' {
        $script:Binding | Should -Not -BeNullOrEmpty
    }

    It 'has a synopsis, a description and at least two examples' {
        $script:Help | Should -Not -BeNullOrEmpty
        ([string]$script:Help.Synopsis).Trim() | Should -Not -BeNullOrEmpty
        ([string]$script:Help.Description).Trim() | Should -Not -BeNullOrEmpty
        @($script:Help.Examples).Count | Should -BeGreaterOrEqual 2
    }

    It 'documents every parameter' {
        $documented = @($script:Help.Parameters.Keys | ForEach-Object { $_.ToUpperInvariant() })
        $missing = @($script:ParamNames | Where-Object { $documented -notcontains $_.ToUpperInvariant() })
        $missing | Should -BeNullOrEmpty
    }

    It 'names the replaced method, the GDAP roles and the partner app permissions in .NOTES' {
        $notes = ([string]$script:Help.Notes) -replace '\s+', ' '
        $notes | Should -Match 'Replaces the original'
        $notes | Should -Match 'Required GDAP roles?:'
        $notes | Should -Match 'Required partner app permissions?:'
    }

    It 'links to the original article and to the toolkit docs' {
        $links = @($script:Help.Links | ForEach-Object { $_.Trim() })
        @($links | Where-Object { $_ -match '^https://gcit\.com\.au/\S+/$' }).Count | Should -BeGreaterOrEqual 1
        @($links | Where-Object { $_ -match '(^|/)docs/\d{2}-[a-z0-9-]+\.md' }).Count | Should -BeGreaterOrEqual 1
    }

    It 'links only to toolkit docs that exist' {
        $docLinks = @($script:Help.Links | ForEach-Object { $_.Trim() } | Where-Object { $_ -notmatch '^https?://' })
        # Either relative to the script (../../docs/05-exchange-access.md) or to the repository root (docs/05-exchange-access.md)
        foreach ($link in $docLinks) {
            $fromScript = Join-Path (Split-Path -Path $Path -Parent) $link
            $fromRoot = Join-Path $script:ScriptsRoot '..' $link
            ((Test-Path -LiteralPath $fromScript) -or (Test-Path -LiteralPath $fromRoot)) | Should -BeTrue -Because "$link should resolve from $Name"
        }
    }

    It 'exposes -TenantId as a string array that accepts pipeline input, or is documented as partner tenant only' {
        if ($script:PartnerTenantOnly -contains $Name) {
            $script:ParamNames | Should -Not -Contain 'TenantId'
            $description = ([string]$script:Help.Description) -replace '\s+', ' '
            $description | Should -Match 'partner tenant'
            $description | Should -Match '-TenantId'
        }
        else {
            $script:ParamNames | Should -Contain 'TenantId'
            $parameter = $script:Ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'TenantId' }
            $parameter.StaticType | Should -Be ([string[]])
            $pipeline = $parameter.Attributes | Where-Object { $_.TypeName.Name -eq 'Parameter' } |
                ForEach-Object { $_.NamedArguments } | Where-Object { $_.ArgumentName -like 'ValueFromPipeline*' }
            $pipeline | Should -Not -BeNullOrEmpty
        }
    }

    It 'has an -OutputPath parameter for CSV output' {
        $script:ParamNames | Should -Contain 'OutputPath'
    }

    It 'uses SupportsShouldProcess and $PSCmdlet.ShouldProcess when it changes anything' {
        $writes = @(Get-WriteCall -Ast $script:Ast | Sort-Object -Unique)
        $hasApply = $script:ParamNames -contains 'Apply'
        if ($writes.Count -gt 0 -or $hasApply) {
            $supports = @($script:Binding.NamedArguments | Where-Object { $_.ArgumentName -eq 'SupportsShouldProcess' })
            $supports.Count | Should -Be 1 -Because "it calls $($writes -join ', ')"
            if ($supports[0].ExpressionOmitted -eq $false) {
                $supports[0].Argument.Extent.Text | Should -Not -Match 'false'
            }
            $script:Content | Should -Match '\$PSCmdlet\.ShouldProcess\(' -Because "it calls $($writes -join ', ')"
        }
    }
}

Describe 'scripts/MAPPING.json' {
    BeforeAll {
        $script:MappingPath = Join-Path $script:ScriptsRoot 'MAPPING.json'
        $script:Mapping = @(Get-Content -LiteralPath $script:MappingPath -Raw | ConvertFrom-Json)
        $script:Standalone = @(Get-ChildItem -Path $script:ScriptsRoot -Directory | ForEach-Object {
                Get-ChildItem -Path $_.FullName -Filter '*.ps1' -File | ForEach-Object { 'scripts/' + $_.Directory.Name + '/' + $_.Name }
            })
    }

    It 'has every field on every entry' {
        $fields = 'article_id', 'title', 'url', 'family', 'status', 'script', 'replaces', 'notes', 'missing_permissions'
        foreach ($entry in $script:Mapping) {
            $names = @($entry.PSObject.Properties.Name)
            foreach ($field in $fields) { $names | Should -Contain $field -Because "article $($entry.article_id)" }
        }
    }

    It 'is sorted by article_id with no duplicates' {
        $ids = @($script:Mapping | ForEach-Object { [int]$_.article_id })
        $ids | Should -Be @($ids | Sort-Object)
        @($ids | Sort-Object -Unique).Count | Should -Be $ids.Count
    }

    It 'uses a known status, and a null script only for obsolete or still-working articles' {
        foreach ($entry in $script:Mapping) {
            $entry.status | Should -BeIn @('broken', 'works-but-outdated', 'obsolete-topic', 'ok')
            if ($entry.status -in 'broken', 'works-but-outdated') {
                $entry.script | Should -Not -BeNullOrEmpty -Because "article $($entry.article_id) is $($entry.status)"
            }
            else {
                $entry.script | Should -BeNullOrEmpty -Because "article $($entry.article_id) is $($entry.status)"
            }
        }
    }

    It 'points only at scripts that exist' {
        foreach ($entry in $script:Mapping | Where-Object script) {
            Test-Path -LiteralPath (Join-Path $script:ScriptsRoot '..' $entry.script) | Should -BeTrue -Because $entry.script
        }
    }

    It 'covers every standalone script' {
        $mapped = @($script:Mapping | Where-Object script | ForEach-Object script)
        @($script:Standalone | Where-Object { $mapped -notcontains $_ }) | Should -BeNullOrEmpty
    }

    It 'links every article to gcit.com.au over HTTPS' {
        foreach ($entry in $script:Mapping) { $entry.url | Should -Match '^https://gcit\.com\.au/\S+/$' }
    }

    It 'has no em or en dashes' {
        # Literal dashes, or their JSON escapes
        Get-Content -LiteralPath $script:MappingPath -Raw | Should -Not -Match ("[$([char]0x2013)$([char]0x2014)]|" + '\\u201[34]')
    }
}

Describe 'scripts/README.md' {
    BeforeAll {
        $script:ReadmePath = Join-Path $script:ScriptsRoot 'README.md'
        $script:Readme = Get-Content -LiteralPath $script:ReadmePath -Raw
    }

    It 'links to every standalone script' {
        $standalone = @(Get-ChildItem -Path $script:ScriptsRoot -Directory | ForEach-Object {
                Get-ChildItem -Path $_.FullName -Filter '*.ps1' -File | ForEach-Object { $_.Directory.Name + '/' + $_.Name }
            })
        foreach ($relative in $standalone) {
            $script:Readme | Should -Match ([regex]::Escape("]($relative)")) -Because $relative
        }
    }

    It 'has only relative links that resolve' {
        $links = [regex]::Matches($script:Readme, '\]\(([^)#\s]+)(#[^)]*)?\)') | ForEach-Object { $_.Groups[1].Value } |
            Where-Object { $_ -notmatch '^[a-z]+://' } | Sort-Object -Unique
        foreach ($link in $links) {
            Test-Path -LiteralPath (Join-Path $script:ScriptsRoot $link) | Should -BeTrue -Because $link
        }
    }

    It 'has PowerShell examples that parse and call scripts that exist with parameters they have' {
        $blocks = [regex]::Matches($script:Readme, '(?ms)^```powershell[ \t]*\r?\n(.*?)^```') | ForEach-Object { $_.Groups[1].Value }
        $blocks | Should -Not -BeNullOrEmpty
        foreach ($code in $blocks) {
            $tokens = $null
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)
            $errors | Should -BeNullOrEmpty
            $calls = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true)
            foreach ($call in $calls) {
                $name = $call.GetCommandName()
                if ($name -notmatch '^\./scripts/') { continue }
                $scriptPath = Join-Path $script:ScriptsRoot '..' $name.Substring(2)
                Test-Path -LiteralPath $scriptPath | Should -BeTrue -Because $name
                $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$null)
                $known = @($scriptAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }) +
                [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
                foreach ($element in $call.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] }) {
                    $known | Should -Contain $element.ParameterName -Because "$name -$($element.ParameterName)"
                }
            }
        }
    }

    It 'has no em or en dashes' {
        $script:Readme | Should -Not -Match "[$([char]0x2013)$([char]0x2014)]"
    }
}
