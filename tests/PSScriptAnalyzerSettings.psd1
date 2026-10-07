@{
    # PSScriptAnalyzer settings for the Pester tests only.
    # The module source (src/) is held to ../PSScriptAnalyzerSettings.psd1 with no exclusions.
    # Test code legitimately breaks some module rules, so these are excluded here, each for a stated reason:
    #
    # PSAvoidGlobalVars
    #   The setup and consent tests share one fake Graph router through $global:MspTest so that dot-sourced
    #   module functions and Pester mocks see the same state. Nothing in src/ uses $global: (checked by
    #   tests/Core.Hygiene.Tests.ps1).
    # PSReviewUnusedParameter, PSShouldProcess, PSUseShouldProcessForStateChangingFunctions, PSUseSingularNouns
    #   Stub functions copy the parameter lists and names of Microsoft cmdlets (Connect-ExchangeOnline,
    #   Set-Secret and others) so mocks can bind. They do nothing, so these rules do not apply.
    # PSUseDeclaredVarsMoreThanAssignments
    #   Variables set in BeforeAll/BeforeEach and read in It blocks are reported as unused because the
    #   analyzer does not follow Pester's scoping.
    # PSUseUsingScopeModifierInNewRunspaces
    #   The listener tests pass values to Start-ThreadJob with -ArgumentList and param(), which the rule
    #   does not recognise.
    # PSPlaceCloseBrace, PSUseCorrectCasing, PSAvoidUsingPositionalParameters, PSUseOutputTypeCorrectly
    #   Layout and style of compact test data (nested hashtables, one-line stubs). Not a correctness issue.

    Severity            = @('Error', 'Warning', 'Information')
    IncludeDefaultRules = $true
    ExcludeRules        = @(
        'PSAvoidGlobalVars'
        'PSReviewUnusedParameter'
        'PSShouldProcess'
        'PSUseShouldProcessForStateChangingFunctions'
        'PSUseSingularNouns'
        'PSUseDeclaredVarsMoreThanAssignments'
        'PSUseUsingScopeModifierInNewRunspaces'
        'PSPlaceCloseBrace'
        'PSUseCorrectCasing'
        'PSAvoidUsingPositionalParameters'
        'PSUseOutputTypeCorrectly'
    )
    Rules               = @{
        PSAvoidUsingConvertToSecureStringWithPlainText = @{ Enable = $true }
        PSAvoidUsingPlainTextForPassword               = @{ Enable = $true }
        PSAvoidUsingWriteHost                          = @{ Enable = $true }
        PSUseCompatibleSyntax                          = @{ Enable = $true; TargetVersions = @('7.4', '7.6') }
    }
}
