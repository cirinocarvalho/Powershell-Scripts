BeforeAll {
    Import-Module -Name (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/OpenDataSync/OpenDataSync.psd1') -Force -ErrorAction Stop

    # Duck typed stand ins for the ADO.NET objects Invoke-SqlInsertBatch drives, so the
    # transaction and disposal behaviour can be verified without a SQL Server instance.
    # These are PSCustomObjects rather than PowerShell classes on purpose: the code under test
    # runs in the module session state, where classes defined in a test file do not resolve.
    function Get-FakeFactory {
        param([switch]$FailOnExecute)

        $state = [pscustomobject]@{
            Opened              = 0
            Committed           = 0
            RolledBack          = 0
            CommandsDisposed    = 0
            TransactionDisposed = 0
            ConnectionDisposed  = 0
            FailOnExecute       = [bool]$FailOnExecute
            Executed            = [System.Collections.Generic.List[object]]::new()
        }

        $factory = [pscustomobject]@{ State = $state }

        Add-Member -InputObject $factory -MemberType ScriptMethod -Name CreateConnection -Value {
            $connection = [pscustomobject]@{ State = $this.State; ConnectionString = $null }

            Add-Member -InputObject $connection -MemberType ScriptMethod -Name Open -Value {
                $this.State.Opened++
            }

            Add-Member -InputObject $connection -MemberType ScriptMethod -Name Dispose -Value {
                $this.State.ConnectionDisposed++
            }

            Add-Member -InputObject $connection -MemberType ScriptMethod -Name BeginTransaction -Value {
                $transaction = [pscustomobject]@{ State = $this.State }
                Add-Member -InputObject $transaction -MemberType ScriptMethod -Name Commit -Value { $this.State.Committed++ }
                Add-Member -InputObject $transaction -MemberType ScriptMethod -Name Rollback -Value { $this.State.RolledBack++ }
                Add-Member -InputObject $transaction -MemberType ScriptMethod -Name Dispose -Value { $this.State.TransactionDisposed++ }
                return $transaction
            }

            Add-Member -InputObject $connection -MemberType ScriptMethod -Name CreateCommand -Value {
                $command = [pscustomobject]@{
                    State          = $this.State
                    Transaction    = $null
                    CommandText    = $null
                    CommandTimeout = 0
                    Parameters     = [System.Collections.Generic.List[object]]::new()
                }

                Add-Member -InputObject $command -MemberType ScriptMethod -Name CreateParameter -Value {
                    return [pscustomobject]@{ ParameterName = $null; Value = $null }
                }

                Add-Member -InputObject $command -MemberType ScriptMethod -Name Dispose -Value {
                    $this.State.CommandsDisposed++
                }

                Add-Member -InputObject $command -MemberType ScriptMethod -Name ExecuteNonQuery -Value {
                    if ($this.State.FailOnExecute) {
                        throw 'simulated deadlock'
                    }

                    $this.State.Executed.Add([pscustomobject]@{
                            CommandText = $this.CommandText
                            Parameters  = $this.Parameters
                        })
                    return 1
                }

                return $command
            }

            return $connection
        }

        return $factory
    }
}

Describe 'Invoke-SqlInsertBatch' {
    BeforeEach {
        $script:Factory = Get-FakeFactory
        Mock -ModuleName OpenDataSync Resolve-SqlClientFactory { $script:Factory }
    }

    It 'opens a connection, executes each row and commits once' {
        $sets = @(
            ([ordered]@{ Id = 'a'; Magnitude = 1.5 }),
            ([ordered]@{ Id = 'b'; Magnitude = 2.5 })
        )

        $affected = InModuleScope OpenDataSync -Parameters @{ sets = $sets } {
            Invoke-SqlInsertBatch -ConnectionString 'fake' -Statement 'INSERT INTO [T] ([Id], [Magnitude]) VALUES (@Id, @Magnitude);' -ParameterSet $sets
        }

        $affected | Should -Be 2
        $script:Factory.State.Opened | Should -Be 1
        $script:Factory.State.Committed | Should -Be 1
        $script:Factory.State.RolledBack | Should -Be 0
        $script:Factory.State.Executed.Count | Should -Be 2
    }

    It 'binds every column as a named parameter rather than inline text' {
        $sets = @([ordered]@{ Id = "x'); DROP TABLE Users--"; Magnitude = 9 })

        $null = InModuleScope OpenDataSync -Parameters @{ sets = $sets } {
            Invoke-SqlInsertBatch -ConnectionString 'fake' -Statement 'INSERT INTO [T] ([Id], [Magnitude]) VALUES (@Id, @Magnitude);' -ParameterSet $sets
        }

        $parameters = $script:Factory.State.Executed[0].Parameters
        $parameters.Count | Should -Be 2
        ($parameters | Where-Object ParameterName -EQ '@Id').Value | Should -Be "x'); DROP TABLE Users--"
        $script:Factory.State.Executed[0].CommandText | Should -Not -Match 'DROP TABLE'
    }

    It 'rolls back and rethrows when a row fails' {
        $script:Factory = Get-FakeFactory -FailOnExecute
        $sets = @([ordered]@{ Id = 'a' })

        $act = {
            InModuleScope OpenDataSync -Parameters @{ sets = $sets } {
                Invoke-SqlInsertBatch -ConnectionString 'fake' -Statement 'INSERT INTO [T] ([Id]) VALUES (@Id);' -ParameterSet $sets
            }
        }

        $act | Should -Throw -ExpectedMessage '*simulated deadlock*'
        $script:Factory.State.RolledBack | Should -Be 1
        $script:Factory.State.Committed | Should -Be 0
    }

    It 'always disposes the connection, transaction and commands' {
        $sets = @([ordered]@{ Id = 'a' })

        $null = InModuleScope OpenDataSync -Parameters @{ sets = $sets } {
            Invoke-SqlInsertBatch -ConnectionString 'fake' -Statement 'INSERT INTO [T] ([Id]) VALUES (@Id);' -ParameterSet $sets
        }

        $script:Factory.State.CommandsDisposed | Should -Be 1
        $script:Factory.State.TransactionDisposed | Should -Be 1
        $script:Factory.State.ConnectionDisposed | Should -Be 1
    }

    It 'disposes the connection even when the run fails' {
        $script:Factory = Get-FakeFactory -FailOnExecute
        $sets = @([ordered]@{ Id = 'a' })

        $act = {
            InModuleScope OpenDataSync -Parameters @{ sets = $sets } {
                Invoke-SqlInsertBatch -ConnectionString 'fake' -Statement 'INSERT INTO [T] ([Id]) VALUES (@Id);' -ParameterSet $sets
            }
        }

        $act | Should -Throw
        $script:Factory.State.ConnectionDisposed | Should -Be 1
        $script:Factory.State.CommandsDisposed | Should -Be 1
    }
}

Describe 'Resolve-SqlClientFactory' {
    It 'gives an actionable error when the SQL client cannot be loaded' {
        $act = {
            InModuleScope OpenDataSync {
                $script:SqlClientFactory = $null
                Mock Import-Module { throw 'module not found' }
                Resolve-SqlClientFactory
            }
        }

        $act | Should -Throw -ExpectedMessage '*Install-Module SqlServer*'
    }
}
