﻿function Get-RemoteRunspace {
    $wc = [System.Management.Automation.Runspaces.WSManConnectionInfo]::new()

    # Use AppVeyor credentials if running in AppVeyor, rather than implicit credentials.
    try
    {
	    $appveyorRemoteCredential = Import-Clixml -Path "$env:TEMP\AppVeyorRemoteCred.xml"
    }
    catch { }
    if ($appveyorRemoteCredential)
    {
        Write-Verbose "Using global AppVeyor credential";
        $wc.Credential = $appveyorRemoteCredential
    }
    else
    {
        Write-Verbose "Using implicit credentials"
    }

    $remoteRunspace = [runspacefactory]::CreateRunspace($host, $wc)
    $remoteRunspace.Open()

    return $remoteRunspace
}

Describe "InvokeOnRunspace method argument error handling" -tags "Feature" {

    BeforeAll {
        $command = [System.Management.Automation.PSCommand]::new()
        $localRunspace = $host.Runspace
    }

    It "Null argument exception should be thrown for null PSCommand argument" {

        try
        {
            [System.Management.Automation.HostUtilities]::InvokeOnRunspace($null, $localRunspace)
            throw "InvokeOnRunspace method did not throw expected PSArgumentNullException exception"
        }
        catch
        {
            $_.FullyQualifiedErrorId | Should Be "PSArgumentNullException"
        }
    }

    It "Null argument exception should be thrown for null Runspace argument" {

        try
        {
            [System.Management.Automation.HostUtilities]::InvokeOnRunspace($command, $null)
            throw "InvokeOnRunspace method did not throw expected PSArgumentNullException exception"
        }
        catch
        {
            $_.FullyQualifiedErrorId | Should Be "PSArgumentNullException"
        }
    }
}

Describe "InvokeOnRunspace method as nested command" -tags "Feature" {

    It "Method should successfully invoke command as nested on busy runspace" {

        $command = [System.Management.Automation.PSCommand]::new()
        $command.AddScript('"Hello!"')
        $currentRunspace = $host.Runspace

        $results = [System.Management.Automation.HostUtilities]::InvokeOnRunspace($command, $currentRunspace)

        $results[0] | Should Be "Hello!"
    }
}

Describe "InvokeOnRunspace method on remote runspace" -tags "Feature" {
    
    BeforeAll {
        $script:skipTest = $true
        It "Get remote runspace for test" -Skip:(!$IsWindows) {
            $script:remoteRunspace = Get-RemoteRunspace
            $script:skipTest = $false
        }
    }

    AfterAll {
        if ($script:remoteRunspace)
        {
            $script:remoteRunspace.Dispose();
        }
    }

    It "Method should successfully invoke command on remote runspace" -Skip:$script:skipTest {

        $command = [System.Management.Automation.PSCommand]::new()
        $command.AddScript('"Hello!"')

        $results = [System.Management.Automation.HostUtilities]::InvokeOnRunspace($command, $script:remoteRunspace)

        $results[0] | Should Be "Hello!"
    }
}
