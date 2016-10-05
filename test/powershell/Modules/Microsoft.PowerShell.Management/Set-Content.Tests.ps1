Describe "Set-Content cmdlet tests" -Tags "CI" {
    BeforeAll {
        $file1 = "file1.txt"
        Setup -File "$file1" -Content $file1
        # if the registry doesn't exist, don't run those tests
        $skipRegistry = ! (test-path hklm:/)
    }
    Context "Set-Content should actually set content" {
        It "should set-Content of testdrive:\$file1" {
            $result=set-content -path testdrive:\$file1 -value "ExpectedContent" -passthru 
            $result| Should be "ExpectedContent"
        }

        It "should return expected string from testdrive:\$file1" {
            $result = get-content -path testdrive:\$file1
            $result | Should BeExactly "ExpectedContent"
        }


        It "should Set-Content to testdrive:\dynamicfile.txt with dynamic parameters" {
            $result=set-content -path testdrive:\dynamicfile.txt -value "ExpectedContent" -passthru
            $result| Should BeExactly "ExpectedContent"
        }

        It "should return expected string from testdrive:\dynamicfile.txt" {
            $result = get-content -path testdrive:\dynamicfile.txt
            $result | Should BeExactly "ExpectedContent"
        }

        It "should remove existing content from testdrive:\$file1 when the -Value is `$null" {
            $AsItWas=get-content testdrive:\$file1
            $AsItWas |Should BeExactly "ExpectedContent"
            set-content -path testdrive:\$file1 -value $null -ea stop
            $AsItIs=get-content testdrive:\$file1
            $AsItIs| Should Not Be $AsItWas
        }

        It "should throw 'ParameterArgumentValidationErrorNullNotAllowed' when -Path is `$null" {
            try {
                set-content -path $null -value "ShouldNotWorkBecausePathIsNull" -ea stop
                Throw "Previous statement unexpectedly succeeded..."
            } 
            catch {
                $_.FullyQualifiedErrorId | Should Be "ParameterArgumentValidationErrorNullNotAllowed,Microsoft.PowerShell.Commands.SetContentCommand"
            }
        }

        It "should throw 'ParameterArgumentValidationErrorNullNotAllowed' when -Path is `$()" {
            try {
                set-content -path $() -value "ShouldNotWorkBecausePathIsInvalid" -ea stop
                Throw "Previous statement unexpectedly succeeded..."
            } 
            catch {
                $_.FullyQualifiedErrorId | Should Be "ParameterArgumentValidationErrorNullNotAllowed,Microsoft.PowerShell.Commands.SetContentCommand"
            }
        }

        It "should throw 'PSNotSupportedException' when you set-content to an unsupported provider" -skip:$skipRegistry {
            try {
                set-content -path HKLM:\\software\\microsoft -value "ShouldNotWorkBecausePathIsUnsupported" -ea stop
                Throw "Previous statement unexpectedly succeeded..."
            } 
            catch {
                $_.FullyQualifiedErrorId | Should Be "NotSupported,Microsoft.PowerShell.Commands.SetContentCommand"
            }
        }
        #[BugId(BugDatabase.WindowsOutOfBandReleases, 9058182)]

        It "should be able to pass multiple [string]`$objects to Set-Content through the pipeline to output a dynamic Path file" {
            "hello","world"|set-content testdrive:\dynamicfile2.txt
            $result=get-content testdrive:\dynamicfile2.txt
            $result.length |Should be 2
            $result[0]     |Should be "hello"
            $result[1]     |Should be "world"
        }
    }
}
