@{
    IncludeDefaultRules = $true

    ExcludeRules        = @(
        # The module deliberately writes operator facing progress to the information stream
        # through Write-AutomationLog, which wraps Write-Information.
        'PSAvoidUsingWriteHost'
    )

    Rules               = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('7.0')
        }

        PSPlaceOpenBrace      = @{
            Enable             = $true
            OnSameLine         = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
        }

        PSUseConsistentIndentation = @{
            Enable          = $true
            Kind            = 'space'
            IndentationSize = 4
        }
    }
}
