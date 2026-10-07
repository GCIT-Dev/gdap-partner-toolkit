@{
    # PSScriptAnalyzer settings for MspGdap.
    # CI fails on any Error, and on the security rules listed in
    # .github/workflows/ci.yml (global variables, Write-Host, plain-text
    # passwords, ConvertTo-SecureString -AsPlainText, missing ShouldProcess).
    # Other warnings are reported and should be fixed, or suppressed in code
    # with [Diagnostics.CodeAnalysis.SuppressMessageAttribute()] and a written
    # justification.

    Severity            = @('Error', 'Warning', 'Information')

    IncludeDefaultRules = $true

    ExcludeRules        = @()

    Rules               = @{
        # Converting a token to SecureString inside the module is allowed only
        # with a SuppressMessageAttribute that explains why.
        PSAvoidUsingConvertToSecureStringWithPlainText = @{
            Enable = $true
        }

        PSAvoidUsingPlainTextForPassword               = @{
            Enable = $true
        }

        PSAvoidUsingUsernameAndPasswordParams          = @{
            Enable = $true
        }

        PSAvoidGlobalVars                              = @{
            Enable = $true
        }

        PSAvoidUsingWriteHost                          = @{
            Enable = $true
        }

        PSUseShouldProcessForStateChangingFunctions    = @{
            Enable = $true
        }

        PSUseApprovedVerbs                             = @{
            Enable = $true
        }

        PSUseCompatibleSyntax                          = @{
            Enable         = $true
            TargetVersions = @('7.4', '7.6')
        }

        PSPlaceOpenBrace                               = @{
            Enable             = $true
            OnSameLine         = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
        }

        PSPlaceCloseBrace                              = @{
            Enable             = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
            NoEmptyLineBefore  = $false
        }

        PSUseConsistentIndentation                     = @{
            Enable              = $true
            IndentationSize     = 4
            PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
            Kind                = 'space'
        }

        PSUseConsistentWhitespace                      = @{
            Enable                          = $true
            CheckInnerBrace                 = $true
            CheckOpenBrace                  = $true
            CheckOpenParen                  = $true
            CheckOperator                   = $false
            CheckPipe                       = $true
            CheckPipeForRedundantWhitespace = $false
            CheckSeparator                  = $true
            CheckParameter                  = $false
        }

        PSUseCorrectCasing                             = @{
            Enable = $true
        }

        PSProvideCommentHelp                           = @{
            Enable                  = $true
            ExportedOnly            = $true
            BlockComment            = $true
            VSCodeSnippetCorrection = $false
            Placement               = 'begin'
        }
    }
}
