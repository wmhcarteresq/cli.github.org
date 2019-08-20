# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

Describe "Validate start of console host" -Tag CI {
    BeforeAll {
        $csharp_source = @'
            using System;
            using System.Runtime.InteropServices;

            public class ScreenReaderTestUtility {
                private const uint SPI_SETSCREENREADER = 0x0047;

                [DllImport("user32.dll", SetLastError = true)]
                [return: MarshalAs(UnmanagedType.Bool)]
                private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

                public static bool ActivateScreenReader() {
                    return SystemParametersInfo(SPI_SETSCREENREADER, 1u, IntPtr.Zero, 0);
                }

                public static bool DeactivateScreenReader() {
                    return SystemParametersInfo(SPI_SETSCREENREADER, 0u, IntPtr.Zero, 0);
                }
            }
'@
        $utilType = "ScreenReaderTestUtility" -as [type]
        if (-not $utilType) {
            $utilType = Add-Type -TypeDefinition $csharp_source -PassThru
        }

        ## Make the screen reader status active.
        $utilType::ActivateScreenReader()
    }

    AfterAll {
        ## Make the screen reader status in-active.
        $utilType::DeactivateScreenReader()
    }

    It "PSReadLine should not be auto-loaded when screen reader status is active" -Skip:(-not $IsWindows) {
        $output = pwsh -noprofile -noexit -c "Get-Module PSReadLine; exit"
        $output.Length | Should -BeExactly 2

        ## The warning message about screen reader should be returned, but the PSReadLine module should not be loaded.
        $output[0] | Should -BeLike "Warning:*'Import-Module PSReadLine'."
        $output[1] | Should -BeExactly [string]::Empty
    }
}
