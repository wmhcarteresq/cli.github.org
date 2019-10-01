# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

Describe 'NullConditionalOperations' -Tags 'CI' {
    BeforeAll {

        $skipTest = -not $EnabledExperimentalFeatures.Contains('PSNullCoalescingOperators')

        if ($skipTest) {
            Write-Verbose "Test Suite Skipped. The test suite requires the experimental feature 'PSNullCoalescingOperators' to be enabled." -Verbose
            $originalDefaultParameterValues = $PSDefaultParameterValues.Clone()
            $PSDefaultParameterValues["it:skip"] = $true
        } else {
            $someGuid = New-Guid
            $typesTests = @(
                @{ name = 'string'; valueToSet = 'hello' }
                @{ name = 'dotnetType'; valueToSet = $someGuid }
                @{ name = 'byte'; valueToSet = [byte]0x94 }
                @{ name = 'intArray'; valueToSet = 1..2 }
                @{ name = 'stringArray'; valueToSet = 'a'..'c' }
                @{ name = 'emptyArray'; valueToSet = @(1, 2, 3) }
            )
        }
    }

    AfterAll {
        if ($skipTest) {
            $global:PSDefaultParameterValues = $originalDefaultParameterValues
        }
    }

    Context "Null conditional assignment operator ??=" {
        It 'Variable doesnot exist' {

            Remove-Variable variableDoesNotExist -ErrorAction SilentlyContinue -Force

            $variableDoesNotExist ??= 1
            $variableDoesNotExist | Should -Be 1

            $variableDoesNotExist ??= 2
            $variableDoesNotExist | Should -Be 1
        }

        It 'Variable exists and is null' {
            $variableDoesNotExist = $null

            $variableDoesNotExist ??= 2
            $variableDoesNotExist | Should -Be 2
        }

        It 'Validate types - <name> can be set' -TestCases $typesTests {
            param ($name, $valueToSet)

            $x = $null
            $x ??= $valueToSet
            $x | Should -Be $valueToSet
        }

        It 'Validate hashtable can be set' {
            $x = $null
            $x ??= @{ 1 = '1' }
            $x.Keys | Should -Be @(1)
        }

        It 'Validate lhs is returned' {
            $x = 100
            $x ??= 200
            $x | Should -Be 100
        }

        It 'Rhs is a cmdlet' {
            $x ??= (Get-Alias -Name 'where')
            $x.Definition | Should -BeExactly 'Where-Object'
        }

        It 'Lhs is DBNull' {
            $x = [System.DBNull]::Value
            $x ??= 200
            $x | Should -Be 200
        }

        It 'Lhs is AutomationNull' {
            $x = [System.Management.Automation.Internal.AutomationNull]::Value
            $x ??= 200
            $x | Should -Be 200
        }

        It 'Lhs is NullString' {
            $x = [NullString]::Value
            $x ??= 200
            $x | Should -Be 200
        }

        It 'Error case' {
            $e = $null
            $null = [System.Management.Automation.Language.Parser]::ParseInput('1 ??= 100', [ref] $null, [ref] $e)
            $e[0].ErrorId | Should -BeExactly 'InvalidLeftHandSide'
        }
    }

    Context 'Null coalesce operator ??' {
        BeforeEach {
            $x = $null
        }

        It 'Variable does not exist' {
            $variableDoesNotExist ?? 100 | Should -Be 100
        }

        It 'Variable exists but is null' {
            $x ?? 100 | Should -Be 100
        }

        It 'Lhs is not null' {
            $x = 100
            $x ?? 200 | Should -Be 100
        }

        It 'Lhs is a non-null constant' {
            1 ?? 2 | Should -Be 1
        }

        It 'Lhs is `$null' {
            $null ?? 'string value' | Should -BeExactly 'string value'
        }

        It 'Check precedence of ?? expression resolution' {
            $x ?? $null ?? 100 | Should -Be 100
            $null ?? $null ?? 100 | Should -Be 100
            $null ?? $null ?? $null | Should -Be $null
            $x ?? 200 ?? $null | Should -Be 200
            $x ?? 200 ?? 300 | Should -Be 200
            100 ?? $x ?? 200 | Should -Be 100
            $null ?? 100 ?? $null ?? 200 | Should -Be 100
        }

        It 'Rhs is a cmdlet' {
            $result = $x ?? (Get-Alias -Name 'where')
            $result.Definition | Should -BeExactly 'Where-Object'
        }

        It 'Lhs is DBNull' {
            $x = [System.DBNull]::Value
            $x ?? 200 | Should -Be 200
        }

        It 'Lhs is AutomationNull' {
            $x = [System.Management.Automation.Internal.AutomationNull]::Value
            $x ??  200 | Should -Be 200
        }

        It 'Lhs is NullString' {
            $x = [NullString]::Value
            $x ?? 200 | Should -Be 200
        }

        It 'Rhs is a get variable expression' {
            $x = [System.DBNull]::Value
            $y = 2
            $x ?? $y | Should -Be 2
        }

        It 'Lhs is a constant' {
            [System.DBNull]::Value ?? 2 | Should -Be 2
        }

        It 'Both are null constants' {
            [System.DBNull]::Value ?? [NullString]::Value | Should -Be ([NullString]::Value)
        }
    }

    Context 'Combined usage of null conditional operators' {

        BeforeAll {
            function GetNull {
                return $null
            }

            function GetHello {
                return "Hello"
            }
        }

        BeforeEach {
            $x = $null
        }

        It '?? and ??= used together' {
            $x ??= 100 ?? 200
            $x | Should -Be 100
        }

        It '?? and ??= chaining' {
            $x ??= $x ?? (GetNull) ?? (GetHello)
            $x | Should -BeExactly 'Hello'
        }
    }
}
