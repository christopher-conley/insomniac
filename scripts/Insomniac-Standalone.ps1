[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)] [string] $TTL,
    [Parameter(Mandatory = $false)] [string] $Activity,
    [Parameter(Mandatory = $false)] [switch] $EXE,
    [Parameter(Mandatory = $false)] [switch] $Help,
    [Parameter(Mandatory = $false)] [string] $Interval,
    [Parameter(Mandatory = $false)] [string] $Key = "{SCROLLLOCK}",
    [Parameter(Mandatory = $false)] [string] $MaxRandomInterval = 237,
    [Parameter(Mandatory = $false)] [string] $MinRandomInterval = 30,
    [Parameter(Mandatory = $false)] [string] $StartAt,
    [Parameter(Mandatory = $false)] [switch] $Version
)

begin {

    $ScriptName = "☕ Insomniac ☕"
    $ScriptVersion = "2024.03.25.2310"
    $Caller = $MyInvocation.MyCommand.Name.ToString()

    function Get-CenteredString {
        param (
            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [array] $String,
            [Parameter(Mandatory = $false)] [int] $Width = $Host.UI.RawUI.BufferSize.Width
        )
        $ReturnObject = $null
        foreach ($Line in $String.Split("`n")) {
            $Line = $Line.PadLeft(([Math]::Max(0, $Width / 2) + [Math]::Floor($Line.Length / 2)))
            $Line = $Line.PadRight($Width)
            $ReturnObject += $Line
        }

        return $ReturnObject
    }

    function Show-VersionAndLicense {
        $LicenseText = "CgBNAEkAVAAgAEwAaQBjAGUAbgBzAGUACgAKAEMAbwBwAHkAcgBpAGcAaAB0ACAAKABjACkAIAAyADAAMgA0ACAAQwBoAHIAaQBzAHQAbwBwAGgAZQByACAAQwBvAG4AbABlAHkACgAKAFAAZQByAG0AaQBzAHMAaQBvAG4AIABpAHMAIABoAGUAcgBlAGIAeQAgAGcAcgBhAG4AdABlAGQALAAgAGYAcgBlAGUAIABvAGYAIABjAGgAYQByAGcAZQAsACAAdABvACAAYQBuAHkAIABwAGUAcgBzAG8AbgAgAG8AYgB0AGEAaQBuAGkAbgBnACAAYQAgAGMAbwBwAHkACgBvAGYAIAB0AGgAaQBzACAAcwBvAGYAdAB3AGEAcgBlACAAYQBuAGQAIABhAHMAcwBvAGMAaQBhAHQAZQBkACAAZABvAGMAdQBtAGUAbgB0AGEAdABpAG8AbgAgAGYAaQBsAGUAcwAgACgAdABoAGUAIAAiAFMAbwBmAHQAdwBhAHIAZQAiACkALAAgAHQAbwAgAGQAZQBhAGwACgBpAG4AIAB0AGgAZQAgAFMAbwBmAHQAdwBhAHIAZQAgAHcAaQB0AGgAbwB1AHQAIAByAGUAcwB0AHIAaQBjAHQAaQBvAG4ALAAgAGkAbgBjAGwAdQBkAGkAbgBnACAAdwBpAHQAaABvAHUAdAAgAGwAaQBtAGkAdABhAHQAaQBvAG4AIAB0AGgAZQAgAHIAaQBnAGgAdABzAAoAdABvACAAdQBzAGUALAAgAGMAbwBwAHkALAAgAG0AbwBkAGkAZgB5ACwAIABtAGUAcgBnAGUALAAgAHAAdQBiAGwAaQBzAGgALAAgAGQAaQBzAHQAcgBpAGIAdQB0AGUALAAgAHMAdQBiAGwAaQBjAGUAbgBzAGUALAAgAGEAbgBkAC8AbwByACAAcwBlAGwAbAAKAGMAbwBwAGkAZQBzACAAbwBmACAAdABoAGUAIABTAG8AZgB0AHcAYQByAGUALAAgAGEAbgBkACAAdABvACAAcABlAHIAbQBpAHQAIABwAGUAcgBzAG8AbgBzACAAdABvACAAdwBoAG8AbQAgAHQAaABlACAAUwBvAGYAdAB3AGEAcgBlACAAaQBzAAoAZgB1AHIAbgBpAHMAaABlAGQAIAB0AG8AIABkAG8AIABzAG8ALAAgAHMAdQBiAGoAZQBjAHQAIAB0AG8AIAB0AGgAZQAgAGYAbwBsAGwAbwB3AGkAbgBnACAAYwBvAG4AZABpAHQAaQBvAG4AcwA6AAoACgBUAGgAZQAgAGEAYgBvAHYAZQAgAGMAbwBwAHkAcgBpAGcAaAB0ACAAbgBvAHQAaQBjAGUAIABhAG4AZAAgAHQAaABpAHMAIABwAGUAcgBtAGkAcwBzAGkAbwBuACAAbgBvAHQAaQBjAGUAIABzAGgAYQBsAGwAIABiAGUAIABpAG4AYwBsAHUAZABlAGQAIABpAG4AIABhAGwAbAAKAGMAbwBwAGkAZQBzACAAbwByACAAcwB1AGIAcwB0AGEAbgB0AGkAYQBsACAAcABvAHIAdABpAG8AbgBzACAAbwBmACAAdABoAGUAIABTAG8AZgB0AHcAYQByAGUALgAKAAoAVABIAEUAIABTAE8ARgBUAFcAQQBSAEUAIABJAFMAIABQAFIATwBWAEkARABFAEQAIAAiAEEAUwAgAEkAUwAiACwAIABXAEkAVABIAE8AVQBUACAAVwBBAFIAUgBBAE4AVABZACAATwBGACAAQQBOAFkAIABLAEkATgBEACwAIABFAFgAUABSAEUAUwBTACAATwBSAAoASQBNAFAATABJAEUARAAsACAASQBOAEMATABVAEQASQBOAEcAIABCAFUAVAAgAE4ATwBUACAATABJAE0ASQBUAEUARAAgAFQATwAgAFQASABFACAAVwBBAFIAUgBBAE4AVABJAEUAUwAgAE8ARgAgAE0ARQBSAEMASABBAE4AVABBAEIASQBMAEkAVABZACwACgBGAEkAVABOAEUAUwBTACAARgBPAFIAIABBACAAUABBAFIAVABJAEMAVQBMAEEAUgAgAFAAVQBSAFAATwBTAEUAIABBAE4ARAAgAE4ATwBOAEkATgBGAFIASQBOAEcARQBNAEUATgBUAC4AIABJAE4AIABOAE8AIABFAFYARQBOAFQAIABTAEgAQQBMAEwAIABUAEgARQAKAEEAVQBUAEgATwBSAFMAIABPAFIAIABDAE8AUABZAFIASQBHAEgAVAAgAEgATwBMAEQARQBSAFMAIABCAEUAIABMAEkAQQBCAEwARQAgAEYATwBSACAAQQBOAFkAIABDAEwAQQBJAE0ALAAgAEQAQQBNAEEARwBFAFMAIABPAFIAIABPAFQASABFAFIACgBMAEkAQQBCAEkATABJAFQAWQAsACAAVwBIAEUAVABIAEUAUgAgAEkATgAgAEEATgAgAEEAQwBUAEkATwBOACAATwBGACAAQwBPAE4AVABSAEEAQwBUACwAIABUAE8AUgBUACAATwBSACAATwBUAEgARQBSAFcASQBTAEUALAAgAEEAUgBJAFMASQBOAEcAIABGAFIATwBNACwACgBPAFUAVAAgAE8ARgAgAE8AUgAgAEkATgAgAEMATwBOAE4ARQBDAFQASQBPAE4AIABXAEkAVABIACAAVABIAEUAIABTAE8ARgBUAFcAQQBSAEUAIABPAFIAIABUAEgARQAgAFUAUwBFACAATwBSACAATwBUAEgARQBSACAARABFAEEATABJAE4ARwBTACAASQBOACAAVABIAEUACgBTAE8ARgBUAFcAQQBSAEUALgAKAA=="
        Write-Host $(Get-CenteredString -String "`n`n$ScriptName")
        Write-Host $(Get-CenteredString -String "Version: $ScriptVersion`n")
        Write-Host $(Get-CenteredString -String "$([System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($LicenseText)))")
    }

    function Test-ObjectIsNullorEmpty {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
            [AllowNull()]
            $TestObject
        )

        begin {
            try {
                $ObjectType = $TestObject.GetType().Name
            }
            catch {
                return
            }
        }

        process {

            switch ($ObjectType) {
                'String' {
                    return ([String]::IsNullOrEmpty($TestObject)) -or $false
                    break
                }

                'DateTime' {
                    try {
                        New-TimeSpan -Start $TestObject -End (Get-Date) -ErrorAction Stop | Out-Null
                        return $false
                    }
                    catch {
                        return $true
                    }
                    break
                }

                'PSCustomObject' {
                    return (@($TestObject.psobject.Properties).Count -eq 0) -or ($null -eq $TestObject)
                    break
                }

                'Hashtable' {
                    return ($TestObject.Count -eq 0) -or ($null -eq $TestObject)
                    break
                }

                Default {
                    Write-MultiStreamMessage -Caller "$Caller" -Messages "Object of type $ObjectType does not have a specific case" -Stream 'debug' -TimeStamp
                    Write-MultiStreamMessage -Caller "$Caller" -Messages "Testing with most likely attributes" -Stream 'debug' -TimeStamp
                    return ($TestObject.Count -eq 0) -or ($TestObject.Length -eq 0) -or ($null -eq $TestObject)
                    break
                }
            }

            ## We should never get here

            Write-Error -Message "Object of type $ObjectType is not supported"
            throw New-Object System.NotSupportedException
        }
        end {

        }
    }

    function Write-MultiStreamMessage {
        param (
            [Parameter(Mandatory = $false)] [string] $Caller,
            [Parameter(Mandatory = $false)] [string] $Stream = "stdout",
            [Parameter(Mandatory = $false)] [switch] $TimeStamp,
            [Parameter(Mandatory = $true)] [array] $Messages
        )

        $Streams = @{
            output  = "Microsoft.PowerShell.Utility\Write-Output"
            stdout  = "Microsoft.PowerShell.Utility\Write-Host"
            stderr  = "Microsoft.PowerShell.Utility\Write-Error"
            verbose = "Microsoft.PowerShell.Utility\Write-Verbose"
            warning = "Microsoft.PowerShell.Utility\Write-Warning"
            info    = "Microsoft.PowerShell.Utility\Write-Information"
            debug   = "Microsoft.PowerShell.Utility\Write-Debug"
        }

        foreach ($Message in $Messages) {
            if ($TimeStamp) {
                $FormattedTimestamp = "$((Get-Date -Format 'o'))|"
                if ($Caller.Length -gt 0) {
                    $Caller += "()"
                    $Message = "$FormattedTimestamp ${Caller}: $Message"
                }
                else {
                    $Message = "$FormattedTimestamp $Message"
                }
            }
            else {
                if ($Caller.Length -gt 0) {
                    $Caller += "()"
                    $Message = "${Caller}: $Message"
                }
            }
            Invoke-Expression -Command "$($Streams.$Stream) '$Message'"
        }
    }

    function Write-ProgressBar {
        param(
            [Parameter(Mandatory = $true)] [hashtable] $SettingsObject
        )

        $Message = $SettingsObject.LoadingMessage
        $ConsoleWidth = ($Host.UI.RawUI.BufferSize.Width - $Message.Length - 10)

        while ($SettingsObject.Timer.Elapsed.Seconds -lt $SettingsObject.Interval) {
            $Progress = [math]::Round(($SettingsObject.Timer.Elapsed.Seconds / $SettingsObject.Interval) * 100)
            $ProgressBar = ([math]::Round(($Progress / 100) * $ConsoleWidth))
            $ProgressBarString = ("#" * $ProgressBar) + '🚀'
            $ProgressString = "{0} {1}% [{2}{3}]" -f $Message, $Progress, $ProgressBarString, (' ' * ($ConsoleWidth - $ProgressBar))

            Write-Host -ForegroundColor Yellow $ProgressString -NoNewline
            Write-Host -ForegroundColor Yellow "`r" -NoNewline
        }

        $ProgressString = "{0} {1}% [{2}]" -f $Message, "100", ("#" * ($ConsoleWidth - 1))
        Write-Host -ForegroundColor Yellow $ProgressString -NoNewline
        Write-Information -MessageData "`n" -InformationAction 'Continue'
        Clear-Host
    }
    function Start-ScriptCleanup {
        param(
            [Parameter(Mandatory = $false)] [hashtable] $SettingsObject
        )
        $SettingsObject.Timer.Stop()
        if ($SettingsObject.PSMajorVersion -ge 7) {
            Stop-Job $SettingsObject.SendKeyJob.Id -ErrorAction SilentlyContinue | Out-Null
            Remove-Job $SettingsObject.SendKeyJob.Id -ErrorAction SilentlyContinue | Out-Null
        }
        Write-MultiStreamMessage -Timestamp -Stream 'verbose' -Caller "$Caller" -Messages "Exiting, $($SettingsObject.TotalIterations) total iteration(s)...`n"
        exit $LASTEXITCODE
    }

    function C8H10N4O2 {
        <#
.SYNOPSIS
    Prevents a login session from timing out due to inactivity.

.DESCRIPTION
    Prevents a login session from timing out due to inactivity by sending a keystroke at a random or specified interval.

    Default settings are:
        TTL:                    Date and time of script start time, plus 69 years
        Interval:               Randomly generated
        Key:                    Scroll Lock
        Activity:               Random loading message
        MinRandomInterval:      30
        MaxRandomInterval:      237
        StartAt:                Current time

    Random loading messages taken from:
    https://gist.github.com/meain/6440b706a97d2dd71574769517e7ed32


.PARAMETER Activity
    Descriptive text to the left of the progress bar.
    Defaults to a random loading message.

.PARAMETER EXE
    When specified, the script will use an internal function to display a progress bar and loading message instead of PowerShell's native Write-Progress function.
    This flag exists to support building the script as a Windows executable with ps2exe.
    Executables built with ps2exe do not support the Write-Progress cmdlet, so this flag is required to display a progress bar and loading message.

.PARAMETER Interval
    The time interval at which to send the keypress.
    Interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    Interval defaults to a random number of seconds between 30 and 237.

.PARAMETER Key
    The keyboard key to send.
    When specifying the Key parameter, enclose special keys like Backspace, Space, Enter, etc. with braces, e.g.:
    {BACKSPACE}
    A full list of special keycodes is available at:
    https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys
    Key defaults to {SCROLLLOCK}

.PARAMETER MaxRandomInterval
    When using a random interval (default setting if -Interval is not specified), the maximum number of seconds between progress bar and loading message resets.
    Defaults to 237 seconds.

.PARAMETER MinRandomInterval
    When using a random interval (default setting if -Interval is not specified), the minimum number of seconds between progress bar and loading message resets.
    Defaults to 30 seconds.

.PARAMETER StartAt
    The date, time, or both at which the script should start sending keypresses.
    StartAt is specified as a date/time string, like "2024-02-27 11:01:32 AM". You may use any date/time format that PowerShell can parse, or even directly pass a .NET [datetime] object.
    StartAt defaults to the current time if not specified.

.PARAMETER TTL
    How long the script should run, specified as a time string, like "7h2m15s". Spaces are accepted, so "7h 2m 15s" is also a valid TTL.
    All three time fields are not required, so "2h", "10m2s", "15m", and "37s" are all valid TTLs.
    TTL defaults to 69 years in the future if not specified.

.EXAMPLE
    .\Insomniac-Standalone.ps1

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop.

.EXAMPLE
    .\Insomniac-Standalone.ps1 7h2m15s

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval for the specified period of time, or until stopped with CTRL+C, or by closing the program. The random interval changes with each loop. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -TTL 10m40s

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval for the specified period of time, or until stopped with CTRL+C, or by closing the program. The random interval changes with each loop. This functions exactly the same as Example #2, with the difference being the timeout is being explicitly set instead of being inferred from commandline input.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -Interval 51

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 51 seconds until stopped with CTRL+C or by closing the program. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 15m -Interval 10s

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 10 seconds for a duration of 15 minutes, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -TTL 37s -Interval 2s

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 2 seconds for a duration of 37 seconds, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -TTL 30m -Key "{BACKSPACE}"

    This example will immediately send a Backspace character, then periodically send a Backspace character at a random interval for a duration of 30 minutes, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -Key "{F15}" -Interval 120 -Activity "Loading..."

    This example will immediately send a F15 character, then periodically send a F15 character every 120 seconds for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -Key "{NUMLOCK}" -MaxRandomInterval 300

    This example will immediately send a Numlock character, then periodically send a Numlock character at a random interval never exceeding 300 seconds for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The MaxRandomInterval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond maximum interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -Key "{SCROLLLOCK}" -MinRandomInterval 20s -MaxRandomInterval 157

    This example will immediately send a Scrolllock character, then periodically send a Scrolllock character at a random interval at least every 20 seconds, but never exceeding 157 seconds, for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The MaxRandomInterval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond maximum interval is also valid.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -StartAt "2024-02-27 11:01:32 AM"

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until February 27th, 2024 at 11:01:32 AM.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -StartAt "07/05/2024"

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until July 5th, 2024 at Midnight.

.EXAMPLE
    .\Insomniac-Standalone.ps1 -StartAt "10:30AM"

    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until 10:30AM (current day).

.INPUTS
    None

.OUTPUTS
    None

.LINK
    https://github.com/christopher-conley/insomniac

.NOTES
    Author         : Christopher Conley <chris@unnx.net>
    Prerequisite   : PowerShell Version 5.1 or higher
    License        : MIT

    Copyright © 2024 Christopher Conley
#>
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [Parameter(Mandatory = $false, ValueFromPipeline = $true)] [string] $TTL,
            [Parameter(Mandatory = $false)] [string] $Activity,
            [Parameter(Mandatory = $false)] [switch] $EXE,
            [Parameter(Mandatory = $false)] [string] $Interval,
            [Parameter(Mandatory = $false)] [string] $Key = "{SCROLLLOCK}",
            [Parameter(Mandatory = $false)] [string] $MaxRandomInterval = 237,
            [Parameter(Mandatory = $false)] [string] $MinRandomInterval = 30
        )

        begin {

            Clear-Host
            $Caller = $MyInvocation.MyCommand.Name.ToString()
            $ToggleKeys = @(
                "{CAPSLOCK}",
                "{INSERT}",
                "{INS}",
                "{NUMLOCK}",
                "{SCROLLLOCK}"
            )

            $SendKeyBlock = @'
 -ScriptBlock {
    param($Settings)
    $Key = $Settings.KeyToSend
    $TimeToComplete = $Settings.Interval
    $TimerSecondsElapsed = $Settings.Timer.Elapsed.Seconds
    $WShell = New-Object -ComObject WScript.Shell
    while ($TimerSecondsElapsed -lt $TimeToComplete) {
        $WShell.SendKeys("$Key")
        if ($Settings.Toggle) {
            Start-Sleep -Milliseconds 200
            $WShell.SendKeys("$Key")
        }
        Start-Sleep $TimeToComplete
        if ($Settings.OriginalInterval -eq 0) {
            $TimeToComplete = Get-Random -Minimum $Settings.MinRandomInterval -Maximum $Settings.MaxRandomInterval
        }
    }
    Clear-Variable WShell -ErrorAction SilentlyContinue | Out-Null
    Remove-Variable WShell -ErrorAction SilentlyContinue | Out-Null
}
'@

            $Settings = @{
                StartTime            = (Get-Date)
                OriginalInterval     = [float](($Interval -replace '\W', '') -replace '[a-zA-Z]', '')
                Interval             = [float](($Interval -replace '\W', '') -replace '[a-zA-Z]', '')
                MinRandomInterval    = [float](($MinRandomInterval -replace '\W', '') -replace '[a-zA-Z]', '')
                MaxRandomInterval    = [float](($MaxRandomInterval -replace '\W', '') -replace '[a-zA-Z]', '')
                WillExpire           = $false
                Timer                = [Diagnostics.Stopwatch]::new()
                LoadingMessage       = $Activity
                RandomLoadingMessage = $false
                Warnings             = @()
                Toggle               = $false
                PSMajorVersion       = $PSVersionTable.PSVersion.Major
                JobStartCmd          = $null
                SendKeyJob           = $null
                SendKeyBlock         = $SendKeyBlock
                KeyToSend            = $Key
                WShellObject         = New-Object -ComObject WScript.Shell
                EXE                  = $EXE
                TotalIterations      = 0
                TTL                  = @{
                    RawTTL       = $TTL
                    TTLHours     = 0
                    TTLMinutes   = 0
                    TTLSeconds   = 0
                    TTLFormatted = $null
                    ExpireAt     = $null
                }
            }

            if ($ToggleKeys -contains $Settings.KeyToSend) {
                $Settings.Toggle = $true
            }

            if ($null -eq $Activity -or $Activity.Length -eq 0) {
                $Settings.LoadingMessage = Get-RandomLoadingMessage
                $Settings.RandomLoadingMessage = $true
            }

            if ($Settings.MinRandomInterval -gt $Settings.MaxRandomInterval) {
                $Settings.Warnings += "Minimum interval is greater than maximum interval, setting minimum interval to a random number between 1 and $($Settings.MaxRandomInterval)"
                $Settings.MinRandomInterval = Get-Random -Minimum 1 -Maximum $Settings.MaxRandomInterval
                $Settings.Warnings += "Minimum interval is now: $($Settings.MinRandomInterval)"
            }

            if ($Settings.TTL.RawTTL.Length -ne 0) {
                $Settings.WillExpire = $true

                $TTLisValid = $Settings.TTL.RawTTL -imatch '(?<Time1Str>\d+\s*[hms])\s*(?<Time2Str>\d+\s*[ms])\s*(?<Time3Str>\d+\s*[s])|(?<Time1Str>\d+\s*[hms])\s*(?<Time2Str>\d+\s*[ms])|(?<Time1Str>\d+\s*[hms])'
                if ($TTLisValid) {
                    $TTLMatches = $Matches
                    $TTLMatches.0 = $null

                    $HasHours = $TTLMatches.Values | Where-Object { $_ -match 'h' }
                    $HasMinutes = $TTLMatches.Values | Where-Object { $_ -match 'm' }
                    $HasSeconds = $TTLMatches.Values | Where-Object { $_ -match 's' }

                    if ($HasHours) {
                        $Settings.TTL.TTLHours = [string](($HasHours -replace '\W', '') -replace '[a-zA-Z]', '')
                    }
                    else {
                        $Settings.TTL.TTLHours = "00"
                    }
                    if ($HasMinutes) {

                        $Settings.TTL.TTLMinutes = [string](($HasMinutes -replace '\W', '') -replace '[a-zA-Z]', '')
                    }
                    else {
                        $Settings.TTL.TTLMinutes = "00"
                    }
                    if ($HasSeconds) {

                        $Settings.TTL.TTLSeconds = [string](($HasSeconds -replace '\W', '') -replace '[a-zA-Z]', '')
                    }
                    else {
                        $Settings.TTL.TTLSeconds = "00"
                    }

                    $Settings.TTL.TTLFormatted = "$($Settings.TTL.TTLHours):$($Settings.TTL.TTLMinutes):$($Settings.TTL.TTLSeconds)"

                    try {
                        $ExpireTimeSpan = ([TimeSpan]"$($Settings.TTL.TTLFormatted)").TotalHours
                        $Settings.TTL.ExpireAt = $Settings.StartTime.AddHours($ExpireTimeSpan)
                    }
                    catch {
                        $Settings.Warnings += "Error creating timespan. TTL will be set to 8 hours."
                        $Settings.TTL.ExpireAt = $Settings.StartTime.AddHours(8)
                    }
                }
                else {
                    $Settings.Warnings += "The TTL provided ($TTL) is not a valid time descriptor. TTL will be set to 8 hours."
                    $Settings.TTL.ExpireAt = $Settings.StartTime.AddHours(8)
                }
            }

            if ($Interval) {
                if ($Settings.Interval -eq 0) {
                    $Settings.Warnings += "The interval provided ($Interval) is not valid. Interval will be random."
                }
            }

            $Settings.Timer = [Diagnostics.Stopwatch]::new()
        }

        process {

            try {

                if (!$Settings.WillExpire) {
                    ## Nice
                    $LoopEnd = (Get-Date).AddYears(69)
                }
                else {
                    $LoopEnd = $Settings.TTL.ExpireAt
                }

                if ($Settings.Warnings.Count -gt 0) {
                    Write-MultiStreamMessage -Timestamp -Stream 'warning' -Caller "$Caller" -Messages $Settings.Warnings
                    Write-Information -MessageData "`n" -InformationAction 'Continue'
                    Write-MultiStreamMessage -Timestamp -Stream 'warning' -Caller "$Caller" "Warning messages will clear in 5 seconds."
                    Start-Sleep 5
                    $LoopEnd = $LoopEnd.AddSeconds(5)
                    Clear-Host
                }

                if ($Interval) {
                    if ($Settings.Interval -eq 0) {
                        $Settings.Interval = Get-Random -Minimum $Settings.MinRandomInterval -Maximum $Settings.MaxRandomInterval
                    }
                }
                else {
                    $Settings.Interval = Get-Random -Minimum $Settings.MinRandomInterval -Maximum $Settings.MaxRandomInterval
                }

                if ($Settings.PSMajorVersion -ge 7) {
                    $Settings.JobStartCmd = 'Start-ThreadJob'
                    $ExpressionCmd = $Settings.JobStartCmd + $Settings.SendKeyBlock + ' -ArgumentList ($Settings)'
                    $Settings.SendKeyJob = (Invoke-Expression -Command $ExpressionCmd)
                }

                ## Old versions of PowerShell don't send the key when using Start-Job for some reason, and I just
                ## don't care enough to figure out why right now. Leaving this else statement here for when I
                ## inevitably forget why I left this here and have some time to figure it out.
                #
                #else {
                #    $Settings.JobStartCmd = 'Start-Job'
                #    $ExpressionCmd = $Settings.JobStartCmd + $Settings.SendKeyBlock + ' -ArgumentList ($Settings)'
                #    $Settings.SendKeyJob = (Invoke-Expression -Command $ExpressionCmd)
                #}

                while (([datetime]::Now) -lt $LoopEnd) {
                    $i++
                    $Settings.Timer.Reset()
                    $Settings.Timer.Start()

                    if ($Interval) {
                        if ($Settings.Interval -eq 0) {
                            $Settings.Interval = Get-Random -Minimum $Settings.MinRandomInterval -Maximum $Settings.MaxRandomInterval
                        }
                    }
                    else {
                        $Settings.Interval = Get-Random -Minimum $Settings.MinRandomInterval -Maximum $Settings.MaxRandomInterval
                    }

                    ## Remove this block after figuring out why when using Start-Job, PowerShell 5
                    ## doesn't send the keypress. It's probably something stupid.
                    #
                    if ($Settings.PSMajorVersion -eq 5) {
                        $Settings.WShellObject.SendKeys($Settings.KeyToSend)
                        if ($Settings.Toggle) {
                            Start-Sleep -Milliseconds 200
                            $Settings.WShellObject.SendKeys($Settings.KeyToSend)
                        }
                    }
                    ##

                    while ($Settings.Timer.Elapsed.Seconds -lt $Settings.Interval) {

                        if (([datetime]::Now) -ge $LoopEnd) {
                            Write-MultiStreamMessage -Timestamp -Stream 'verbose' -Caller "$Caller" -Messages "TTL hit, breaking...`n"
                            break
                        }

                        if ($Settings.EXE) {
                            Write-ProgressBar -SettingsObject $Settings
                        }

                        else {
                            $Completed = ($Settings.Timer.Elapsed.Seconds / $Settings.Interval) * 100

                            $ProgressSplat = @{
                                Activity         = $Settings.LoadingMessage
                                Status           = " "
                                PercentComplete  = $Completed
                                CurrentOperation = $CurrentOperation
                                SecondsRemaining = ($Settings.Interval - $Settings.Timer.Elapsed.Seconds)
                            }
                            Write-Progress @ProgressSplat
                            Start-Sleep -Seconds ($Settings.Interval * 0.25)
                        }
                    }
                    if ($Settings.RandomLoadingMessage) {
                        $Settings.LoadingMessage = Get-RandomLoadingMessage
                    }
                    $Settings.TotalIterations++
                }
            }
            finally {
                Start-ScriptCleanup -SettingsObject $Settings
            }
        }
        end {

        }
    }

    if ($Version) {
        Show-VersionAndLicense
        exit 0
    }

    if ($Help) {
        Get-Help C8H10N4O2 -Full
        exit 0
    }

    switch -Regex ($TTL) {
        "^--help$|^-help$|^help$|^-h$|^\/\?$" {
            Get-Help C8H10N4O2 -Full
            exit 0
        }
        "^--version$|^-version$|^version$|^-v$" {
            Show-VersionAndLicense
            exit 0
        }
    }

    function Get-RandomLoadingMessage {

        $LoadingMessages = @("Reticulating splines...",
            "Generating witty dialog...",
            "Swapping time and space...",
            "Spinning violently around the y-axis...",
            "Tokenizing real life...",
            "Bending the spoon...",
            "Filtering morale...",
            "Don't think of purple hippos...",
            "We need a new fuse...",
            "Have a good day.",
            "Upgrading Windows, your PC will restart several times. Sit back and relax.",
            "640K ought to be enough for anybody",
            "The architects are still drafting",
            "The bits are breeding",
            "We're building the buildings as fast as we can",
            "Would you prefer chicken, steak, or tofu?",
            "(Pay no attention to the man behind the curtain)",
            "...and enjoy the elevator music...",
            "Please wait while the little elves draw your map",
            "Don't worry - a few bits tried to escape, but we caught them",
            "Would you like fries with that?",
            "Checking the gravitational constant in your locale...",
            "Go ahead -- hold your breath!",
            "...at least you're not on hold...",
            "Hum something loud while others stare",
            "You're not in Kansas any more",
            "The server is powered by a lemon and two electrodes.",
            "Please wait while a larger software vendor in Seattle takes over the world",
            "We're testing your patience",
            "As if you had any other choice",
            "Follow the white rabbit",
            "Why don't you order a sandwich?",
            "While the satellite moves into position",
            "keep calm and npm install",
            "The bits are flowing slowly today",
            "Dig on the 'X' for buried treasure... ARRR!",
            "It's still faster than you could draw it",
            "The last time I tried this the monkey didn't survive. Let's hope it works better this time.",
            "I should have had a V8 this morning.",
            "My other loading screen is much faster.",
            "Testing on Timmy... We're going to need another Timmy.",
            "Reconfoobling energymotron...",
            "(Insert quarter)",
            "Are we there yet?",
            "Have you lost weight?",
            "Just count to 10",
            "Why so serious?",
            "It's not you. It's me.",
            "Counting backwards from Infinity",
            "Don't panic...",
            "Embiggening Prototypes",
            "Do not run! We are your friends!",
            "Do you come here often?",
            "Warning: Don't set yourself on fire.",
            "We're making you a cookie.",
            "Creating time-loop inversion field",
            "Spinning the wheel of fortune...",
            "Loading the enchanted bunny...",
            "Computing chance of success",
            "I'm sorry Dave, I can't do that.",
            "Looking for exact change",
            "All your web browser are belong to us",
            "All I really need is a kilobit.",
            "I feel like I'm supposed to be loading something...",
            "What do you call 8 Hobbits? A Hobbyte.",
            "Should have used a compiled language...",
            "Is this Windows?",
            "Adjusting flux capacitor...",
            "Please wait until the sloth starts moving.",
            "Don't break your screen yet!",
            "I swear it's almost done.",
            "Let's take a mindfulness minute...",
            "Unicorns are at the end of this road, I promise.",
            "Listening for the sound of one hand clapping...",
            "Keeping all the 1's and removing all the 0's...",
            "Putting the icing on the cake. The cake is not a lie...",
            "Cleaning off the cobwebs...",
            "Making sure all the i's have dots...",
            "We need more dilithium crystals",
            "Where did all the internets go",
            "Connecting Neurotoxin Storage Tank...",
            "Granting wishes...",
            "Time flies when you're having fun.",
            "Get some coffee and come back in ten minutes..",
            "Spinning the hamster…",
            "99 bottles of beer on the wall..",
            "Stay awhile and listen..",
            "Be careful not to step in the git-gui",
            "You shall not pass! yet..",
            "Load it and they will come",
            "Convincing AI not to turn evil..",
            "There is no spoon. Because we are not done loading it",
            "Your left thumb points to the right and your right thumb points to the left.",
            "How did you get here?",
            "Wait, do you smell something burning?",
            "Computing the secret to life, the universe, and everything.",
            "When nothing is going right, go left!!...",
            "I love my job only when I'm on vacation...",
            "I'm not lazy, I'm just relaxed!!",
            "Never steal. The government hates competition....",
            "Why are they called apartments if they are all stuck together?",
            "Life is Short - Talk Fast!",
            "Optimism - is a lack of information.....",
            "Save water and shower together",
            "Whenever I find the key to success, someone changes the lock.",
            "Sometimes I think war is God's way of teaching us geography.",
            "I've got problem for your solution…..",
            "Where there's a will, there's a relative.",
            "Adults are just kids with money.",
            "I think I am, therefore, I am. I think.",
            "A kiss is like a fight, with mouths.",
            "Coffee, Chocolate, Men. The richer the better!",
            "git happens",
            "May the forks be with you",
            "A commit a day keeps the mobs away",
            "This is not a joke, it's a commit.",
            "Constructing additional pylons...",
            "Roping some seaturtles...",
            "Locating Jebediah Kerman...",
            "We are not liable for any broken screens as a result of waiting.",
            "Hello IT, have you tried turning it off and on again?",
            "If you type Google into Google you can break the internet",
            "Well, this is embarrassing.",
            "What is the airspeed velocity of an unladen swallow?",
            "Hello, IT... Have you tried forcing an unexpected reboot?",
            "They just toss us away like yesterday's jam.",
            "They're fairly regular, the beatings, yes. I'd say we're on a bi-weekly beating.",
            "The Elders of the Internet would never stand for it.",
            "Space is invisible mind dust, and stars are but wishes.",
            "Didn't know paint dried so quickly.",
            "Everything sounds the same",
            "I'm going to walk the dog",
            "I didn't choose the engineering life. The engineering life chose me.",
            "Dividing by zero...",
            "Spawn more Overlord!",
            "If I'm not back in five minutes, just wait longer.",
            "Some days, you just can't get rid of a bug!",
            "We're going to need a bigger boat.",
            "Web developers do it with <style>",
            "I need to git pull --my-life-together",
            "Java developers never RIP. They just get Garbage Collected.",
            "Cracking military-grade encryption...",
            "Simulating traveling salesman...",
            "Proving P=NP...",
            "Entangling superstrings...",
            "Twiddling thumbs...",
            "Searching for plot device...",
            "Trying to sort in O(n)...",
            "Laughing at your pictures- I mean, loading...",
            "Sending data to NS- I mean, our servers.",
            "Looking for sense of humour, please hold on.",
            "Please wait while the intern refills his coffee.",
            "A different error message? Finally, some progress!",
            "Hold on while we wrap up our git together...sorry",
            "Please hold on as we reheat our coffee",
            "Kindly hold on as we convert this bug to a feature...",
            "Kindly hold on as our intern quits vim...",
            "Winter is coming...",
            "Installing dependencies",
            "Switching to the latest JS framework...",
            "Distracted by cat gifs",
            "Finding someone to hold my beer",
            "BRB, working on my side project",
            "@todo Insert witty loading message",
            "Let's hope it's worth the wait",
            "Aw, snap! Not..",
            "Ordering 1s and 0s...",
            "Updating dependencies...",
            "Whatever you do, don't look behind you...",
            "Please wait... Consulting the manual...",
            "It is dark. You're likely to be eaten by a grue.",
            "Loading funny message...",
            "It's 10:00pm. Do you know where your children are?",
            "Waiting for Daenerys to say all her titles...",
            "Feel free to spin in your chair",
            "What the what?",
            "format C: ...",
            "Forget you saw that password I just typed into the IM ...",
            "What's under there?",
            "Your computer has a virus, its name is Windows!",
            "Go ahead, hold your breath and do an ironman plank till loading complete",
            "Bored of slow loading spinner, buy more RAM!",
            "Help, I'm trapped in a loader!",
            "What is the difference between a hippo and a zippo? One is really heavy, the other is a little lighter",
            "Please wait, while we purge the Decepticons for you. Yes, You can thanks us later!",
            "Mining some bitcoins...",
            "Downloading more RAM..",
            "Updating to Windows Vista...",
            "Deleting System32 folder",
            "Hiding all ;'s in your code",
            "Alt-F4 speeds things up.",
            "Initializing the initializer...",
            "When was the last time you dusted around here?",
            "Optimizing the optimizer...",
            "Last call for the data bus! All aboard!",
            "Running swag sticker detection...",
            "Never let a computer know you're in a hurry.",
            "A computer will do what you tell it to do, but that may be much different from what you had in mind.",
            "Some things man was never meant to know. For everything else, there's Google.",
            "Unix is user-friendly. It's just very selective about who its friends are.",
            "Shovelling coal into the server",
            "Pushing pixels...",
            "How about this weather, eh?",
            "Building a wall...",
            "Everything in this universe is either a potato or not a potato",
            "The severity of your issue is always lower than you expected.",
            "Updating Updater...",
            "Downloading Downloader...",
            "Debugging Debugger...",
            "Reading Terms and Conditions for you.",
            "Digested cookies being baked again.",
            "Live long and prosper.",
            "There is no cow level, but there's a goat one!",
            "Running with scissors...",
            "Definitely not a virus...",
            "You may call me Steve.",
            "You seem like a nice person...",
            "Coffee at my place, tommorow at 10AM - don't be late!",
            "Work, work...",
            "Patience! This is difficult, you know...",
            "Discovering new ways of making you wait...",
            "Your time is very important to us. Please wait while we ignore you...",
            "Time flies like an arrow; fruit flies like a banana",
            "Two men walked into a bar; the third ducked...",
            "Sooooo... Have you seen my vacation photos yet?",
            "Sorry we are busy catching em' all, we're done soon",
            "TODO: Insert elevator music",
            "Still faster than Windows update",
            "Composer hack: Waiting for reqs to be fetched is less frustrating if you add -vvv to your command.",
            "Please wait while the minions do their work",
            "Grabbing extra minions",
            "Doing the heavy lifting",
            "We're working very hard... really",
            "Waking up the minions",
            "You are number 2843684714 in the queue",
            "Please wait while we serve other customers...",
            "Our premium plan is faster",
            "Feeding unicorns...",
            "Rupturing the subspace barrier",
            "Creating an anti-time reaction",
            "Converging tachyon pulses",
            "Bypassing control of the matter-antimatter integrator",
            "Adjusting the dilithium crystal converter assembly",
            "Reversing the shield polarity",
            "Disrupting warp fields with an inverse graviton burst",
            "Up, Up, Down, Down, Left, Right, Left, Right, B, A.",
            "Do you like my loading animation? I made it myself",
            "Whoah, look at it go!",
            "No, I'm awake. I was just resting my eyes.",
            "One mississippi, two mississippi...",
            "Don't panic... AHHHHH!",
            "Ensuring Gnomes are still short.",
            "Baking ice cream...")

        $RandomNum = Get-Random -Minimum 0 -Maximum $($LoadingMessages.Length - 1)
        return $LoadingMessages[$RandomNum]
    }

    $SendKeysSpecialChars = @(
        '+',
        '^',
        '%',
        '~',
        '(',
        ')',
        '[',
        ']'
    )

    if ($null -ne $Key -or $Key.Length -ne 0) {

        $KeyToSend = $Key
        if ($KeyToSend[0] -eq '{' -and $KeyToSend[$KeyToSend.Length - 1] -eq '}') {
            $IsNonPrintableKey = $true
            $KeyToSend = $KeyToSend[1..($KeyToSend.Length - 2)] -join ''
        }
        $KeyToSend = $KeyToSend -replace "({|})", '{$1}'
        foreach ($SpecialCharacter in $SendKeysSpecialChars) {
            $KeyToSend = $KeyToSend -replace "\$SpecialCharacter", "{$SpecialCharacter}"
        }
        if ($IsNonPrintableKey) {
            $KeyToSend = '{' + "$KeyToSend" + '}'
        }
    }

    $SettingsSplat = @{
        TTL               = $TTL
        Interval          = $Interval
        Key               = $KeyToSend
        Activity          = $Activity
        MinRandomInterval = $MinRandomInterval
        MaxRandomInterval = $MaxRandomInterval
        EXE               = $EXE
    }

    $ProcessName = (Get-Process -PID $PID).ProcessName

    if ($ProcessName -notmatch "^pwsh$|^powershell$") {
        ## Running as an executable, not a script
        $SettingsSplat.EXE = $true
    }

}

process {

    if (!(Test-ObjectIsNullorEmpty $StartAt)) {
        try {
            $ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
            $CurrentTime = Get-Date
            $BeginRunningAt = Get-Date -Date $StartAt -ErrorAction Stop
            $BeginRunningAtFormatted = (Get-Date "$BeginRunningAt" -UFormat '%Y-%m-%d %r' -ErrorAction Stop)

            while ($CurrentTime -lt $BeginRunningAt) {
                Clear-Host
                Write-Information -MessageData "Preparing to launch... 🚀`n" -InformationAction 'Continue'

                $TimeRemaining = New-TimeSpan -Start $CurrentTime -End $BeginRunningAt
                $LaunchMessage = "Preparations will take approximately: $([Math]::Round($TimeRemaining.TotalSeconds)) seconds (until $BEginRunningAtFormatted)..."
                $LaunchString = "{0}{1}" -f $LaunchMessage, (' ' * ($ConsoleWidth - $LaunchMessage.Length))
                Write-Host $LaunchString -NoNewline

                $SleepTime = $TimeRemaining.TotalSeconds * 0.1
                Start-Sleep -Seconds $SleepTime
                $CurrentTime = [datetime]::Now
            }
        }
        catch {
            $ErrorObject = $_
            $ErrorMessages = @(
                "Error type '$($ErrorObject.CategoryInfo.Category)' occurred while parsing -StartAt time. Please provide a valid time format. Valid time formats are any time string accepted by the Get-Date cmdlet, like:",
                "'MM/dd/yyyy HH:mm:ss AM', 'yyyy/MM/dd HH:mm PM', 'MM/dd/yyyy HH:mm AM', 'yyyy/MM/dd', 'HH:mm:ss AM', 'HH:mm PM'",
                "You can also pass a .NET [datetime] object directly to -StartAt. For example:",
                '-StartAt $((Get-Date).AddMinutes(30)',
                "-StartAt (Get-Date '11:30AM')",
                "Exiting..."
            )
            foreach ($ErrorMessage in $ErrorMessages) {
                Write-Error -Message $ErrorMessage
            }
            exit 11
        }
    }

    C8H10N4O2 @SettingsSplat
}

end {

}


















# SIG # Begin signature block
# MIIopQYJKoZIhvcNAQcCoIIoljCCKJICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDlwgkNH7nnYd4N
# ginSxVSYzdRt3pzrnVxeClqeyCoJtaCCDaUwgga5MIIEoaADAgECAhEAmaOACiZV
# O2Wr3G6EprPqOTANBgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNV
# BAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBD
# ZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQg
# TmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjEL
# MAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEk
# MCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34
# vuTmflN4mSAfgLKTvggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSo
# McbvIKck6+hI1shsylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+
# njvMk2xqbNTIPsnWtw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHI
# ohzuC8KNqfcYf7Z4/iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRda
# CtVD2bSlbfsq7BiqljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837
# Q4QOZgYqVWQ4x6cM7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8m
# FSkcSkAExzd4prHwYjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI
# 5XDYO962WZayx7ACFf5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5bou
# fUnq1UiYPIAHlezf4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplK
# ShBSND36E/ENVv8urPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n8
# 58JSqPECAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10
# XUwA23ufoHTKsW73PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbR
# Og79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8E
# KTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsG
# AQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVt
# LmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0
# bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6
# Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCia
# U58Q7EP89DttyZqGYn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L9
# 4C9LGF3vjzzH8Jq3iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rG
# DxLUUAsO0eqeLNhLVsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f
# 8pXdd3x2mbJCKKtl2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0Cd
# Y9rNLqyA3ahe8WlxVWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B
# 7eMcZNhpn8zJ+6MTyE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloM
# U+vUBfSouCReZwSLo8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1t
# VxTpV2Na3PR7nxYVlPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2
# DaGgK2W1eEJfo2qyrBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNj
# t2ywzOr+tKyEVAotnyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCr
# W602NCmvO1nm+/80nLy5r0AZvCQxaQ4wggbkMIIEzKADAgECAhA7U5aXFldstcsh
# Wwg7IMKaMA0GCSqGSIb3DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNp
# Z25pbmcgMjAyMSBDQTAeFw0yNDAzMjUxNjU3NDVaFw0yNTAzMjUxNjU3NDRaMIGJ
# MQswCQYDVQQGEwJVUzERMA8GA1UECAwIVmlyZ2luaWExEzARBgNVBAcMCkFsZXhh
# bmRyaWExHjAcBgNVBAoMFU9wZW4gU291cmNlIERldmVsb3BlcjEyMDAGA1UEAwwp
# T3BlbiBTb3VyY2UgRGV2ZWxvcGVyLCBDaHJpc3RvcGhlciBDb25sZXkwggIiMA0G
# CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0ePATYBR3lyqet0e5n7lSGxpIcBDc
# wbmsHV7JhnRBEGfCrbWIJSqrGnRtthyxQH0xb55GFxW+U2dOjx74HclR7GfiBl5B
# fwZtUSI4vaWkRw/FkHNeEc+L9j4w1exnUev0WaL+yNSTuOLF7q9vygY8TC38eC2b
# pCJ9KlczloGWj4CM/dwPmBjZTavru5+Hie7M4RyxE8gYCgC+6mbvhzvtxhdxPXFy
# L2/6fVRAHZIBJc6ta5K0NYXc//Rfb2f7c2y+twThr6KnAOW+nhUpoZ9mU8JfnD53
# T11q5GxCYjc56PUrtOdjV5wHee2EPuuI8xRR4eBL5FZl7D5cvWgGWpX43g7+NF2G
# WNLab25P0z6h1ogfXB4fW3Rx/BU4vww6LU5tk9Z9mopE+V5lWV7BBSb+KgyoTTCI
# p9afwEU3YgLCgfVw6eVrYAEfpEeMGQY97m/sv6NvSaQcTZo6wM4CgrjmFFFlDNey
# xOszB9Nm6zRBfb9wOmKiY4P4DwnMKc4PZx7JB8shW05IygWCifN9U0eHe3qLoSvh
# NYqSGSflvVAxa0aJASPaYCffR8dDPeumWMRacAn0uFPJtxydoI4X732XcqUblgvC
# C8mmqZo1SvgkW/vJrAByQmpURLEmabwgMe1l2L6nX0cvhbovtjIUWuSfGGNKN2fk
# Wa32gxFTlweSGQIDAQABo4IBeDCCAXQwDAYDVR0TAQH/BAIwADA9BgNVHR8ENjA0
# MDKgMKAuhixodHRwOi8vY2NzY2EyMDIxLmNybC5jZXJ0dW0ucGwvY2NzY2EyMDIx
# LmNybDBzBggrBgEFBQcBAQRnMGUwLAYIKwYBBQUHMAGGIGh0dHA6Ly9jY3NjYTIw
# MjEub2NzcC1jZXJ0dW0uY29tMDUGCCsGAQUFBzAChilodHRwOi8vcmVwb3NpdG9y
# eS5jZXJ0dW0ucGwvY2NzY2EyMDIxLmNlcjAfBgNVHSMEGDAWgBTddF1MANt7n6B0
# yrFu9zzAMsBwzTAdBgNVHQ4EFgQUQ4COF3+Ix+Ok+JBZ3PoIdagMkPIwSwYDVR0g
# BEQwQjAIBgZngQwBBAEwNgYLKoRoAYb2dwIFAQQwJzAlBggrBgEFBQcCARYZaHR0
# cHM6Ly93d3cuY2VydHVtLnBsL0NQUzATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNV
# HQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAC1Oi6ByZPHnRP4fTQWfYNcf
# hfKKekheLYXW1adY2TZzChwTe5l0PAn15Nt+1JkouusgOfIKYltzFtO3PaOdjZ2H
# N0diw6twJRsyHs/B/cGSHQ1taJdmb6JZ6rZUsck46QkzuSuceBXFQywNji3T6Th7
# nAJpS21oVgdolQKdalwAEXO3P8AKgGGbjMWam6dHML28cbTFZKRoC6A4nNeXmTCq
# D2mT68l/Ybmic8BOBdMQ8X3yScDIOvoHcn9or2jH5lfLY8JkUto5wMYCIXg8id5J
# CvWu+D3mnEYFjoymf/Z5OaOE60w8w199T/c8d0EknuKaJCroBdPlvDjjqKHfLc8G
# BY5Aa6eYukWxCJ3Mo609lHrb33q2gVpadhRsCoGLOboAD7/KQcDih0kg2FWgLhWs
# Qw2QOiNMPVL7kvkKuhU5NoFDjWuT7sUx8vb7cyDfhiK4dSmt6ChmfO11dIfQwkFk
# I/AAjuKuX+NtQNDfLcAyIBbzMFGg5Inoubz//F3RHqPpz/W4oebThCHpMIF8JOVO
# DNtODSXa5/IrCYdsTJwJE0rTc0sC35w6+c6VLVtFlBdEWZU/jmkG7qEK2H3FuDLO
# KCa5dYMaHBMh0pxUpKKF7slvKyk3s3gLluNg1WGoQVHip+IUooka+OYtA95tSfEc
# EIuGjXJneFe9xX3eJDNfMYIaVjCCGlICAQEwajBWMQswCQYDVQQGEwJQTDEhMB8G
# A1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0g
# Q29kZSBTaWduaW5nIDIwMjEgQ0ECEDtTlpcWV2y1yyFbCDsgwpowDQYJYIZIAWUD
# BAIBBQCgfDAQBgorBgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQx
# IgQg9XSRI1jxP7lGlDWa3E9fSA++Z80LqNdsOqEd3tYYelYwDQYJKoZIhvcNAQEB
# BQAEggIABhlOSeVYNAsR4z4G4tYA7ZLI5KDLr008mijXVdHvA4wpo1WTu8HL61w4
# noBx9oZ3jfJS7pdKJvOoOQtxkFRnS3HgNVaKDmuw0t38P0iaalLH/IQFA7gIqDZ3
# FGD1On4MGItf3/09UuXbhFmu8Qfv4cIn2iMQIOU1lAdzwc8dCueOeT6G2WzEOI7n
# zbfgvUsxudqQYChxOc5Y18k7CO09AAhthnX+HsBT7apPQUL88gi5dOVUU2n/tKFE
# o0PEHEBMo5GOPOB28QIxgiHkCsnqers5z/P4Mq7DjIYZNMpKQVWDZo9JcDrltoGY
# oyfhzWg/vUDvpK93Tpv3SvaHo0prmepo9CwhWqy8pdIGlKJ91G1AeahZPpcpM45O
# oaq0JRFVkDN9us80njlTK6gRNrnxoninvYoF0UVuAWMQ4dUhRT7QtCc+ksbREum5
# mIuPo6y/y7VB/dx0T8aWlm+j4QG1x4VuzYeZcBSEoVeDYmPt08hMdOfSJteWRSRO
# H4hqhLuIKEsp905tIW7NFYpG9/vIn2sME9HeSfc/UCP++buiR66lCfAdjRQiJkBt
# WF2iJWB2cOhV+rgKc5zC/LRGftE+P6sLtO25N1lUrh8zTQZOp4vOQp5ssAMqSbfC
# emVLKREo+dQ6yIclmEci+ybEq9jWSJ/JBmV14AoDOGWK05EmiK2hghc/MIIXOwYK
# KwYBBAGCNwMDATGCFyswghcnBgkqhkiG9w0BBwKgghcYMIIXFAIBAzEPMA0GCWCG
# SAFlAwQCAQUAMHcGCyqGSIb3DQEJEAEEoGgEZjBkAgEBBglghkgBhv1sBwEwMTAN
# BglghkgBZQMEAgEFAAQgsDxHiEQ1NjTpJj82WTE7dHXincJPKsLpKfxWr8MDaa0C
# EG1CPQnft+Wa5Azbf0N6kfwYDzIwMjQwMzI3MTQwOTEzWqCCEwkwggbCMIIEqqAD
# AgECAhAFRK/zlJ0IOaa/2z9f5WEWMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYT
# AlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQg
# VHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjMw
# NzE0MDAwMDAwWhcNMzQxMDEzMjM1OTU5WjBIMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFtcCAy
# MDIzMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAo1NFhx2DjlusPlSz
# I+DPn9fl0uddoQ4J3C9Io5d6OyqcZ9xiFVjBqZMRp82qsmrdECmKHmJjadNYnDVx
# vzqX65RQjxwg6seaOy+WZuNp52n+W8PWKyAcwZeUtKVQgfLPywemMGjKg0La/H8J
# JJSkghraarrYO8pd3hkYhftF6g1hbJ3+cV7EBpo88MUueQ8bZlLjyNY+X9pD04T1
# 0Mf2SC1eRXWWdf7dEKEbg8G45lKVtUfXeCk5a+B4WZfjRCtK1ZXO7wgX6oJkTf8j
# 48qG7rSkIWRw69XloNpjsy7pBe6q9iT1HbybHLK3X9/w7nZ9MZllR1WdSiQvrCuX
# vp/k/XtzPjLuUjT71Lvr1KAsNJvj3m5kGQc3AZEPHLVRzapMZoOIaGK7vEEbeBlt
# 5NkP4FhB+9ixLOFRr7StFQYU6mIIE9NpHnxkTZ0P387RXoyqq1AVybPKvNfEO2hE
# o6U7Qv1zfe7dCv95NBB+plwKWEwAPoVpdceDZNZ1zY8SdlalJPrXxGshuugfNJgv
# OuprAbD3+yqG7HtSOKmYCaFxsmxxrz64b5bV4RAT/mFHCoz+8LbH1cfebCTwv0KC
# yqBxPZySkwS0aXAnDU+3tTbRyV8IpHCj7ArxES5k4MsiK8rxKBMhSVF+BmbTO776
# 65E42FEHypS34lCh8zrTioPLQHsCAwEAAaOCAYswggGHMA4GA1UdDwEB/wQEAwIH
# gDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZ
# MBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2
# mi91jGogj57IbzAdBgNVHQ4EFgQUpbbvE+fvzdBkodVWqWUxo97V40kwWgYDVR0f
# BFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1
# c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNybDCBkAYIKwYBBQUH
# AQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBY
# BggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAgRrW3qCptZgXvHCNT4o8aJzYJf/LLOTN6l0ikuyMIgKpuM+A
# qNnn48XtJoKKcS8Y3U623mzX4WCcK+3tPUiOuGu6fF29wmE3aEl3o+uQqhLXJ4Xz
# jh6S2sJAOJ9dyKAuJXglnSoFeoQpmLZXeY/bJlYrsPOnvTcM2Jh2T1a5UsK2nTip
# gedtQVyMadG5K8TGe8+c+njikxp2oml101DkRBK+IA2eqUTQ+OVJdwhaIcW0z5iV
# GlS6ubzBaRm6zxbygzc0brBBJt3eWpdPM43UjXd9dUWhpVgmagNF3tlQtVCMr1a9
# TMXhRsUo063nQwBw3syYnhmJA+rUkTfvTVLzyWAhxFZH7doRS4wyw4jmWOK22z75
# X7BC1o/jF5HRqsBV44a/rCcsQdCaM0qoNtS5cpZ+l3k4SF/Kwtw9Mt911jZnWon4
# 9qfH5U81PAC9vpwqbHkB3NpE5jreODsHXjlY9HxzMVWggBHLFAx+rrz+pOt5Zapo
# 1iLKO+uagjVXKBbLafIymrLS2Dq4sUaGa7oX/cR3bBVsrquvczroSUa31X/MtjjA
# 2Owc9bahuEMs305MfR5ocMB3CtQC4Fxguyj/OOVSWtasFyIjTvTs0xf7UGv/B3cf
# cZdEQcm4RtNsMnxYL2dHZeUbc7aZ+WssBkbvQR7w8F/g29mtkIBEr4AQQYowggau
# MIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAe
# Fw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcw
# FQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3Rl
# ZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9Ge
# TKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0
# hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZl
# jZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAsh
# aG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVY
# TXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1
# biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP65x9abJTyUpURK1h0QCir
# c0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+
# DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA
# +bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42Pg
# puE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzS
# M7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU
# uhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6
# mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsG
# AQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAE
# GTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1Z
# jsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8d
# B+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVp
# P0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp8
# 76i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2
# nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3
# ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQ
# txMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc
# 4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+Y
# AN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZ
# vAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQr
# H4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIFjTCCBHWgAwIBAgIQDpsY
# jvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQw
# IgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAw
# MDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhE
# aWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57
# G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9o
# k3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFh
# mzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463J
# T17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFw
# q1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yh
# Tzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU
# 75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LV
# jHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJ
# bOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/Y+whX8Qg
# UWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IB
# OjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6
# mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/
# BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3Au
# ZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4
# oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJv
# b3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBw
# oL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0
# E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtD
# IeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlU
# sLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCwrFig
# DkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwY
# w02fc7cBqZ9Xql4o4rmUMYIDdjCCA3ICAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUG
# A1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQg
# RzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhAFRK/zlJ0IOaa/2z9f
# 5WEWMA0GCWCGSAFlAwQCAQUAoIHRMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAcBgkqhkiG9w0BCQUxDxcNMjQwMzI3MTQwOTEzWjArBgsqhkiG9w0BCRACDDEc
# MBowGDAWBBRm8CsywsLJD4JdzqqKycZPGZzPQDAvBgkqhkiG9w0BCQQxIgQgsVCc
# lNjecox1PuVA2Qmcn50n8Ttm1z4peEi2n/waZRIwNwYLKoZIhvcNAQkQAi8xKDAm
# MCQwIgQg0vbkbe10IszR1EBXaEE2b4KK2lWarjMWr00amtQMeCgwDQYJKoZIhvcN
# AQEBBQAEggIAXI/WDovMauyJP1YOpBSfDKuKusK7fQ/9w8oUpsp/B701fqpKx6ww
# HdStBRu/ZKSumS9kAopftJH75ooGWLt0kO+43DF5poqBi8kdqP92gCQIxhEke4QG
# E480Re5kVbBtHjOdp5fKNAfMmawfl5n2TNWJokiIcJqdFP/M1Gz+6/7dEWGGPq81
# ppuQj220bw6EOk/pBLGmURpN3wOP6I1kbNh0J9WyKbrirm8h37+itIpDRfq+GRLF
# l4hBdQR4LQM9GkmNHFpBC+rtCWoen7r7UnfrtedfS4KMHkkiJKUC+c41U4B7/pVl
# u/ddwev89mCIdimorb2eig29WoFNyi3icK/ITZ+uTUIXQTMNqPBZ6Qr/0Wskaz+d
# nuMTHVeOejqfNBHx7h1emDjz1JqYh/HY6ej0qeyVQ7toFduECbqliHFK3YU/ulaa
# /O1j8lpCSSjz5PLtWMsAcepHldlSii4tabACz0w3FpVc74MsbgELwdzthZ5hnki5
# SGdkpXlwAodRF33RcJ8l9lOYaX6ev124jxzSBMYwoorxUqhmYoI7oxHYW6L16XdA
# WtoCvx/3cVu7gQqCUR7p/2msSEvCZ4482/0572tgIRSpPte7kbOZOTg6yv66KspC
# wNI/GLaVzHjZEw8KamheH00ZpObDLJdhwRtHcwDGsrdp0lerC9vs5Fw=
# SIG # End signature block
