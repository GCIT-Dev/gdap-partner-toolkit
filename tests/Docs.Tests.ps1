#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Checks that every PowerShell example in the README and the guides still matches the module:
# - every fenced powershell block parses
# - every MspGdap command it calls exists and is exported
# - every named parameter it passes (including keys of a splatted hashtable) exists on that command,
#   by full name, unambiguous prefix or alias
# - every *-Msp* command mentioned anywhere in the text exists
# Nothing here runs the examples.

BeforeDiscovery {
    $repoRoot = Join-Path $PSScriptRoot '..'
    $docFiles = @(Get-ChildItem -Path (Join-Path $repoRoot 'docs') -Filter '*.md' -File) + @(Get-Item -Path (Join-Path $repoRoot 'README.md'))
    $script:DocBlocks = foreach ($file in $docFiles) {
        $text = Get-Content -LiteralPath $file.FullName -Raw
        $index = 0
        foreach ($match in [regex]::Matches($text, '(?ms)^```(?:powershell|pwsh|ps1|ps)[ \t]*\r?\n(.*?)^```')) {
            $index++
            $line = ($text.Substring(0, $match.Index) -split "`n").Count
            @{ File = $file.Name; Block = $index; Line = $line; Code = $match.Groups[1].Value }
        }
    }
    $script:DocFiles = foreach ($file in $docFiles) { @{ File = $file.Name; Path = $file.FullName } }
}

BeforeAll {
    $manifest = Join-Path $PSScriptRoot '..' 'src' 'MspGdap' 'MspGdap.psd1'
    Import-Module $manifest -Force
    $script:Commands = @{}
    foreach ($command in Get-Command -Module MspGdap) { $script:Commands[$command.Name] = $command }
    $script:Common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters

    function Get-DocParameterProblem {
        param([string]$Code)
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$tokens, [ref]$errors)
        foreach ($parseError in $errors) { "parse error: $($parseError.Message)" }
        $calls = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($call in $calls) {
            $name = $call.GetCommandName()
            if (-not $name -or $name -notmatch '-Msp') { continue }
            if (-not $script:Commands.ContainsKey($name)) { "unknown command: $name"; continue }
            $command = $script:Commands[$name]
            $used = New-Object System.Collections.Generic.List[string]
            foreach ($element in $call.CommandElements) {
                if ($element -is [System.Management.Automation.Language.CommandParameterAst]) { $used.Add($element.ParameterName) }
                if ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                    $variable = $element.VariablePath.UserPath
                    $assignment = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq "`$$variable" }, $true) | Select-Object -Last 1
                    if ($assignment) {
                        $table = $assignment.Right.FindAll({ param($node) $node -is [System.Management.Automation.Language.HashtableAst] }, $true) | Select-Object -First 1
                        if ($table) { foreach ($pair in $table.KeyValuePairs) { $used.Add($pair.Item1.Extent.Text.Trim("'`"")) } }
                    }
                }
            }
            foreach ($parameter in $used) {
                if ($script:Common -contains $parameter) { continue }
                $exact = $command.Parameters.ContainsKey($parameter) -or @($command.Parameters.Values | Where-Object { $_.Aliases -contains $parameter }).Count -gt 0
                $prefix = @($command.Parameters.Keys | Where-Object { $_ -like "$parameter*" })
                if (-not $exact -and $prefix.Count -ne 1) { "$name has no parameter -$parameter" }
            }
        }
    }
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Documentation example <File> block <Block> (line <Line>)' -ForEach $script:DocBlocks {
    It 'parses and only uses existing MspGdap commands and parameters' {
        @(Get-DocParameterProblem -Code $Code) | Should -BeNullOrEmpty
    }
}

Describe 'Documentation text <File>' -ForEach $script:DocFiles {
    It 'only mentions MspGdap commands that exist' {
        $text = Get-Content -LiteralPath $Path -Raw
        $unknown = foreach ($match in [regex]::Matches($text, '\b[A-Z][a-z]+-Msp[A-Za-z]*\b(?!\*)')) {
            # Wildcard mentions such as Connect-Msp* are skipped.
            if (-not $script:Commands.ContainsKey($match.Value)) { $match.Value }
        }
        @($unknown | Select-Object -Unique) | Should -BeNullOrEmpty
    }
    It 'uses no em dashes or en dashes' {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Not -Match "[$([char]0x2014)$([char]0x2013)]"
    }
}

Describe 'Documentation checker self-test' {
    It 'catches an unknown parameter, an unknown command, a bad splat key and a parse error' {
        @(Get-DocParameterProblem -Code "Grant-MspPartnerAppConsent -TenantId 'x' -NoSuchSwitch") | Should -Match 'NoSuchSwitch'
        @(Get-DocParameterProblem -Code 'Remove-MspEverything -TenantId x') | Should -Match 'unknown command'
        @(Get-DocParameterProblem -Code "`$p = @{ TenantId = 'x'; Bogus = 1 }`nTest-MspPartnerAppConsent @p") | Should -Match 'Bogus'
        @(Get-DocParameterProblem -Code 'Get-MspCustomer -Name (') | Should -Match 'parse error'
        @(Get-DocParameterProblem -Code "Grant-MspPartnerAppConsent -CustomerTenantId 'x' -WhatIf") | Should -BeNullOrEmpty
    }
}
