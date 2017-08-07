Describe "CredSSP cmdlet tests" -Tags 'Feature','RequireAdminOnWindows' {

    BeforeAll {
        $powershell = Join-Path $PSHOME "powershell"
        $notEnglish = $false

        $originalDefaultParameterValues = $PSDefaultParameterValues.Clone()
        if ( ! $IsWindows )
        {
            $PSDefaultParameterValues["it:skip"] = $true
        }
        else 
        {
            if ([System.Globalization.CultureInfo]::CurrentCulture.Name -ne "en-US")
            {
                $notEnglish = $true
            }
        }
    }
    
    AfterAll {
        $global:PSDefaultParameterValues = $originalDefaultParameterValues
    }

    BeforeEach {
        $errtxt = "$testdrive/error.txt"
        Remove-Item $errtxt -Force -ErrorAction SilentlyContinue
        $donefile = "$testdrive/done"
        Remove-Item $donefile -Force -ErrorAction SilentlyContinue
    }

    It "Error returned if invalid parameters: <description>" -TestCases @(
        @{params=@{Role="Client"};Description="Client role, no DelegateComputer"},
        @{params=@{Role="Server";DelegateComputer="."};Description="Server role w/ DelegateComputer"}
    ) {
        param ($params)
        { Enable-WSManCredSSP @params } | ShouldBeErrorId "System.InvalidOperationException,Microsoft.WSMan.Management.EnableWSManCredSSPCommand"
    }

    It "Enable-WSManCredSSP works: <description>" -Skip:($NotEnglish) -TestCases @(
        @{params=@{Role="Client";DelegateComputer="*"};description="client"},
        @{params=@{Role="Server"};description="server"}
    ) {
        param ($params)
        $c = Enable-WSManCredSSP @params -Force
        $c.CredSSP | Should Be $true

        $c = Get-WSManCredSSP
        if ($params.Role -eq "Client")
        {
            $c[0] | Should Match "The machine is configured to allow delegating fresh credentials to the following target\(s\):wsman/\*"
        }
        else
        {
            $c[1] | Should Match "This computer is configured to receive credentials from a remote client computer"
        }
    }

    It "Disable-WSManCredSSP works: <role>" -Skip:($NotEnglish) -TestCases @(
        @{Role="Client"},
        @{Role="Server"}
    ) {
        param ($role)
        Disable-WSManCredSSP -Role $role | Should BeNullOrEmpty

        $c = Get-WSManCredSSP
        if ($role -eq "Client")
        {
            $c[0] | Should Match "The machine is not configured to allow delegating fresh credentials."
        }
        else
        {
            $c[1] | Should Match "This computer is not configured to receive credentials from a remote client computer"
        }
    }

    It "Call cmdlet as API" {
        $credssp = [Microsoft.WSMan.Management.EnableWSManCredSSPCommand]::new()
        $credssp.Role = "Client"
        $credssp.Role | Should BeExactly "Client"
        $credssp.DelegateComputer = "foo", "bar"
        $credssp.DelegateComputer -join ',' | Should Be "foo,bar"
        $credssp.Force = $true
        $credssp.Force | Should Be $true

        $credssp = [Microsoft.WSMan.Management.DisableWSManCredSSPCommand]::new()
        $credssp.Role = "Server"
        $credssp.Role | Should BeExactly "Server"
    }

    It "Error returned if runas non-admin: <cmdline>" -TestCases @(
        @{cmdline = "Enable-WSManCredSSP -Role Server -Force"; cmd = "EnableWSManCredSSPCommand"},
        @{cmdline = "Disable-WSManCredSSP -Role Server"; cmd = "DisableWSManCredSSPCommand"},
        @{cmdline = "Get-WSManCredSSP"; cmd = "GetWSmanCredSSPCommand"}
    ) {
        param ($cmdline, $cmd)

        runas.exe /trustlevel:0x20000 "$powershell -nop -c try { $cmdline } catch { `$_.FullyQualifiedErrorId | Out-File $errtxt }; New-Item -Type File -Path $donefile"
        $startTime = Get-Date
        while (((Get-Date) - $startTime).TotalSeconds -lt 5 -and -not (Test-Path "$donefile"))
        {
            Start-Sleep -Milliseconds 100
        }
        $errtxt | Should Exist
        $err = Get-Content $errtxt
        $err | Should Be "System.InvalidOperationException,Microsoft.WSMan.Management.$cmd"
    }
}
