# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.
Describe "Get-PSDrive" -Tags "CI" {

    It "Should not throw" {
	Get-PSDrive | Should -Not -BeNullOrEmpty
    }

    It "Should have a name and a length property" {
	(Get-PSDrive).Name        | Should -Not -BeNullOrEmpty
	(Get-PSDrive).Root.Length | Should -Not -BeLessThan 1
    }

    # PSAvoidUsingCmdletAliases warnings should be ignored here as the use of aliases is intentional.
    It "Should be able to be called with the gdr alias" {
	{ gdr } | Should -Not -Throw

	gdr | Should -Not -BeNullOrEmpty
    }

    # PSAvoidUsingCmdletAliases warnings should be ignored here as the use of aliases is intentional.
    It "Should be the same output between Get-PSDrive and gdr" {
	$alias  = gdr
	$actual = Get-PSDrive


	$alias | Should -BeExactly $actual
    }

    It "Should return drive info"{
        (Get-PSDrive Env).Name        | Should -BeExactly Env
        (Get-PSDrive Alias).Name      | Should -BeExactly Alias

        if ($IsWindows)
        {
            (Get-PSDrive Cert).Root       | Should -Be \
            (Get-PSDrive C).Provider.Name | Should -BeExactly FileSystem
        }
        else
        {
            (Get-PSDrive /).Provider.Name | Should -BeExactly FileSystem
        }
    }

    It "Should be able to access a drive using the PSProvider switch" {
	(Get-PSDrive -PSProvider FileSystem).Name.Length | Should -BeGreaterThan 0
    }

    It "Should return true that a drive that does not exist"{
	!(Get-PSDrive fake -ErrorAction SilentlyContinue) | Should -BeTrue
	Get-PSDrive fake -ErrorAction SilentlyContinue    | Should -BeNullOrEmpty
    }
    It "Should be able to determine the amount of free space of a drive" {
        $dInfo = Get-PSDrive TESTDRIVE
        $dInfo.Free -ge 0 | Should -BeTrue
    }
    It "Should be able to determine the amount of Used space of a drive" {
        $dInfo = Get-PSDrive TESTDRIVE
        $dInfo.Used -ge 0 | Should -BeTrue
    }
}
