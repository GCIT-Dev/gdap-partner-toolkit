#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Static checks over the module source. These protect the public repo rules:
# no global state, no plain-text secret conversions, no console-only output,
# and every public file is exported.

BeforeDiscovery {
    $script:SourceRoot = Join-Path $PSScriptRoot '..' 'src' 'MspGdap'
    $script:SourceFiles = @(Get-ChildItem -Path $script:SourceRoot -Recurse -Include '*.ps1', '*.psm1' -File | ForEach-Object {
            @{ Name = $_.Name; Path = $_.FullName }
        })
}

BeforeAll {
    $script:SourceRoot = Join-Path $PSScriptRoot '..' 'src' 'MspGdap'
    $script:ManifestPath = Join-Path $script:SourceRoot 'MspGdap.psd1'
}

Describe 'Source hygiene: <Name>' -ForEach $script:SourceFiles {
    BeforeAll {
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $script:Content = Get-Content -LiteralPath $Path -Raw
    }

    It 'does not use $global: variables' {
        $globals = $script:Ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.VariableExpressionAst] -and $node.VariablePath.IsGlobal }, $true)
        @($globals).Count | Should -Be 0
    }

    It 'does not call Write-Host' {
        $calls = $script:Ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Write-Host' }, $true)
        @($calls).Count | Should -Be 0
    }

    It 'does not use ConvertTo-SecureString -AsPlainText' {
        $script:Content | Should -Not -Match 'ConvertTo-SecureString[^\r\n]*-AsPlainText'
    }

    It 'is ASCII only (no em dashes or smart quotes)' {
        $script:Content | Should -Not -Match '[^\x00-\x7F]'
    }
}

Describe 'Module manifest' {
    It 'exports every public function and nothing else' {
        $manifest = Import-PowerShellDataFile -Path $script:ManifestPath
        $public = @(Get-ChildItem -Path (Join-Path $script:SourceRoot 'Public') -Filter '*.ps1' | ForEach-Object BaseName | Sort-Object)
        @($manifest.FunctionsToExport | Sort-Object) | Should -Be $public
    }

    It 'pins PowerShell 7.4 or later and version 0.2.1' {
        $manifest = Import-PowerShellDataFile -Path $script:ManifestPath
        [version]$manifest.PowerShellVersion | Should -BeGreaterOrEqual ([version]'7.4')
        $manifest.ModuleVersion | Should -Be '0.2.1'
    }
}

Describe 'Comment-based help on exported commands' {
    BeforeAll {
        Import-Module $script:ManifestPath -Force
        $script:Common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
    }
    AfterAll { Remove-Module MspGdap -Force -ErrorAction SilentlyContinue }

    It '<_> has a synopsis, an example and a description for every parameter' -ForEach @(
        (Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot '..' 'src' 'MspGdap' 'MspGdap.psd1')).FunctionsToExport
    ) {
        $command = Get-Command -Name $_ -Module MspGdap
        $help = Get-Help -Name $_ -Full
        ([string]$help.Synopsis).Trim() | Should -Not -BeNullOrEmpty
        $help.examples | Should -Not -BeNullOrEmpty
        $undocumented = foreach ($name in $command.Parameters.Keys) {
            if ($script:Common -contains $name) { continue }
            $entry = @($help.parameters.parameter | Where-Object { $_.name -eq $name }) | Select-Object -First 1
            if (-not $entry -or -not (($entry.description | Out-String).Trim())) { $name }
        }
        @($undocumented) | Should -BeNullOrEmpty
    }
}
