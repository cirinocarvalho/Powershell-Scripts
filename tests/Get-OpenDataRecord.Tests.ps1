BeforeAll {
    Import-Module -Name (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1') -Force -ErrorAction Stop
}

Describe 'Get-OpenDataRecord' {
    BeforeAll {
        $script:GeoJson = [pscustomobject]@{
            type     = 'FeatureCollection'
            features = @(
                [pscustomobject]@{ id = 'a'; properties = [pscustomobject]@{ mag = 1.2 } }
                [pscustomobject]@{ id = 'b'; properties = [pscustomobject]@{ mag = 3.4 } }
            )
        }
    }

    It 'returns the records found at -RecordPath' {
        Mock -ModuleName OpenDataSync Invoke-RestMethod { $script:GeoJson }

        $result = Get-OpenDataRecord -Uri 'https://example.test/api' -RecordPath 'features'

        $result.Count | Should -Be 2
        $result[0].id | Should -Be 'a'
    }

    It 'returns the raw response when no -RecordPath is given' {
        Mock -ModuleName OpenDataSync Invoke-RestMethod { $script:GeoJson }

        (Get-OpenDataRecord -Uri 'https://example.test/api').type | Should -Be 'FeatureCollection'
    }

    It 'builds a sorted, URL encoded query string' {
        $script:CapturedUri = $null
        Mock -ModuleName OpenDataSync Invoke-RestMethod {
            $script:CapturedUri = $Uri
            $script:GeoJson
        }

        $null = Get-OpenDataRecord -Uri 'https://example.test/api' -QueryParameter @{
            limit  = 10
            format = 'geo json'
        }

        # AbsoluteUri keeps the percent encoding; ToString() would decode %20 back to a space.
        $script:CapturedUri.AbsoluteUri | Should -Be 'https://example.test/api?format=geo%20json&limit=10'
    }

    It 'appends with & when the URI already has a query string' {
        $script:CapturedUri = $null
        Mock -ModuleName OpenDataSync Invoke-RestMethod {
            $script:CapturedUri = $Uri
            $script:GeoJson
        }

        $null = Get-OpenDataRecord -Uri 'https://example.test/api?v=1' -QueryParameter @{ limit = 5 }

        $script:CapturedUri.AbsoluteUri | Should -Be 'https://example.test/api?v=1&limit=5'
    }

    It 'sends a bearer token when the secret resolves' {
        $script:CapturedHeader = $null
        Mock -ModuleName OpenDataSync Invoke-RestMethod {
            $script:CapturedHeader = $Headers
            $script:GeoJson
        }
        Mock -ModuleName OpenDataSync Get-AutomationSecret { 'super-secret' }

        $null = Get-OpenDataRecord -Uri 'https://example.test/api' -TokenSecretName 'ApiToken'

        $script:CapturedHeader['Authorization'] | Should -Be 'Bearer super-secret'
        $script:CapturedHeader['User-Agent'] | Should -Match 'OpenDataSync'
    }

    It 'continues unauthenticated when the token secret is missing' {
        $script:CapturedHeader = $null
        Mock -ModuleName OpenDataSync Invoke-RestMethod {
            $script:CapturedHeader = $Headers
            $script:GeoJson
        }
        Mock -ModuleName OpenDataSync Get-AutomationSecret { $null }

        $null = Get-OpenDataRecord -Uri 'https://example.test/api' -TokenSecretName 'ApiToken' -WarningAction SilentlyContinue

        $script:CapturedHeader.ContainsKey('Authorization') | Should -BeFalse
    }

    It 'retries transient failures and succeeds' {
        $script:Attempt = 0
        Mock -ModuleName OpenDataSync Start-Sleep { }
        Mock -ModuleName OpenDataSync Invoke-RestMethod {
            $script:Attempt++
            if ($script:Attempt -lt 3) { throw '503 Service Unavailable' }
            $script:GeoJson
        }

        $result = Get-OpenDataRecord -Uri 'https://example.test/api' -RecordPath 'features' `
            -RetryDelaySeconds 0 -WarningAction SilentlyContinue

        $script:Attempt | Should -Be 3
        $result.Count | Should -Be 2
    }

    It 'throws after the retry budget is exhausted' {
        Mock -ModuleName OpenDataSync Start-Sleep { }
        Mock -ModuleName OpenDataSync Invoke-RestMethod { throw 'boom' }

        { Get-OpenDataRecord -Uri 'https://example.test/api' -MaximumRetryCount 2 `
                -RetryDelaySeconds 0 -WarningAction SilentlyContinue } |
            Should -Throw -ExpectedMessage '*failed after 3 attempt(s)*'

        Should -Invoke -ModuleName OpenDataSync Invoke-RestMethod -Times 3 -Exactly
    }

    It 'does not retry when -MaximumRetryCount is 0' {
        Mock -ModuleName OpenDataSync Invoke-RestMethod { throw 'boom' }

        { Get-OpenDataRecord -Uri 'https://example.test/api' -MaximumRetryCount 0 } | Should -Throw

        Should -Invoke -ModuleName OpenDataSync Invoke-RestMethod -Times 1 -Exactly
    }

    It 'throws a clear error for an unknown record path' {
        Mock -ModuleName OpenDataSync Invoke-RestMethod { $script:GeoJson }

        { Get-OpenDataRecord -Uri 'https://example.test/api' -RecordPath 'missing' } |
            Should -Throw -ExpectedMessage "*'missing' was not found*"
    }
}
