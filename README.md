# insomniac
 
Can't run the (excellent software) [Caffeine](https://www.zhornsoftware.co.uk/caffeine/) due to restrictions? Well, here's a poor approximation in PowerShell!

Disclaimer: I have no affliation with Zhorn Software, nor do they endorse this script in any way, shape, form, or fashion.

The purpose of this script is to enable similar functionality in an environment where you might not necessarily have the ability to run arbitrary executables. There **is** an executable available in all the [releaes](https://github.com/christopher-conley/insomniac/releases), but it's not necessary to run the script, it's just another option; the script is completely standalone and has no external dependencies.

Here's the `Get-Help` output for right now until I actually write a README:

```
NAME
    Insomniac-Standalone.ps1
    
SYNOPSIS
    Prevents a login session from timing out due to inactivity.
    
    
SYNTAX
    Insomniac-Standalone.ps1 [[-TTL] <String>] [[-Activity] <String>] [-EXE] [-Help] [[-Interval] <String>] [[-Key] <String>] [[-MaxRandomInterval] <String>] [[-MinRandomInterval] <String>] [[-StartAt] <String>] [-Version] [<CommonParameters>]
    
    
DESCRIPTION
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
    

PARAMETERS
    -TTL <String>
        How long the script should run, specified as a time string, like "7h2m15s". Spaces are accepted, so "7h 2m 15s" is also a valid TTL.
        All three time fields are not required, so "2h", "10m2s", "15m", and "37s" are all valid TTLs.
        TTL defaults to 69 years in the future if not specified.
        
        Required?                    false
        Position?                    1
        Default value                
        Accept pipeline input?       true (ByValue)
        Accept wildcard characters?  false
        
    -Activity <String>
        Descriptive text to the left of the progress bar.
        Defaults to a random loading message.
        
        Required?                    false
        Position?                    2
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -EXE [<SwitchParameter>]
        When specified, the script will use an internal function to display a progress bar and loading message instead of PowerShell's native Write-Progress function.
        This flag exists to support building the script as a Windows executable with ps2exe.
        Executables built with ps2exe do not support the Write-Progress cmdlet, so this flag is required to display a progress bar and loading message.
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Help [<SwitchParameter>]
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Interval <String>
        The time interval at which to send the keypress.
        Interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
        Interval defaults to a random number of seconds between 30 and 237.
        
        Required?                    false
        Position?                    3
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Key <String>
        The keyboard key to send.
        When specifying the Key parameter, enclose special keys like Backspace, Space, Enter, etc. with braces, e.g.:
        {BACKSPACE}
        A full list of special keycodes is available at:
        https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys
        Key defaults to {SCROLLLOCK}
        
        Required?                    false
        Position?                    4
        Default value                {SCROLLLOCK}
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -MaxRandomInterval <String>
        When using a random interval (default setting if -Interval is not specified), the maximum number of seconds between progress bar and loading message resets.
        Defaults to 237 seconds.
        
        Required?                    false
        Position?                    5
        Default value                237
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -MinRandomInterval <String>
        When using a random interval (default setting if -Interval is not specified), the minimum number of seconds between progress bar and loading message resets.
        Defaults to 30 seconds.
        
        Required?                    false
        Position?                    6
        Default value                30
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -StartAt <String>
        The date, time, or both at which the script should start sending keypresses.
        StartAt is specified as a date/time string, like "2024-02-27 11:01:32 AM". You may use any date/time format that PowerShell can parse, or even directly pass a .NET [datetime] object.
        StartAt defaults to the current time if not specified.
        
        Required?                    false
        Position?                    7
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Version [<SwitchParameter>]
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https://go.microsoft.com/fwlink/?LinkID=113216). 
    
INPUTS
    None
    
    
OUTPUTS
    None


NOTES

        Author         : Christopher Conley <chris@unnx.net>
        Prerequisite   : PowerShell Version 5.1 or higher
        License        : MIT
        
        Copyright © 2024 Christopher Conley
    
    -------------------------- EXAMPLE 1 --------------------------
    
    PS > .\Insomniac-Standalone.ps1
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop.
    
    
    
    
    -------------------------- EXAMPLE 2 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 7h2m15s
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval for the specified period of time, or until stopped with CTRL+C, or by closing the program. The random interval changes with each loop. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used.
    
    
    
    
    -------------------------- EXAMPLE 3 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -TTL 10m40s
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval for the specified period of time, or until stopped with CTRL+C, or by closing the program. The random interval changes with each loop. This functions exactly the same as Example #2, with the difference being the timeout is being explicitly set instead of being inferred from commandline input.
    
    
    
    
    -------------------------- EXAMPLE 4 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -Interval 51
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 51 seconds until stopped with CTRL+C or by closing the program. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 5 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 15m -Interval 10s
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 10 seconds for a duration of 15 minutes, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 6 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -TTL 37s -Interval 2s
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character every 2 seconds for a duration of 37 seconds, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 7 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -TTL 30m -Key "{BACKSPACE}"
    
    This example will immediately send a Backspace character, then periodically send a Backspace character at a random interval for a duration of 30 minutes, or until stopped with CTRL+C, or by closing the program. The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 8 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -Key "{F15}" -Interval 120 -Activity "Loading..."
    
    This example will immediately send a F15 character, then periodically send a F15 character every 120 seconds for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The interval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 9 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -Key "{NUMLOCK}" -MaxRandomInterval 300
    
    This example will immediately send a Numlock character, then periodically send a Numlock character at a random interval never exceeding 300 seconds for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The MaxRandomInterval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond maximum interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 10 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -Key "{SCROLLLOCK}" -MinRandomInterval 20s -MaxRandomInterval 157
    
    This example will immediately send a Scrolllock character, then periodically send a Scrolllock character at a random interval at least every 20 seconds, but never exceeding 157 seconds, for a duration of 69 years, or until stopped with CTRL+C, or by closing the program. The message displayed adjacent to the progress bar will be "Loading...". The time string does not require all three hours/minutes/seconds fields, so a string of "7h" would run for 7 hours, "30m10s" would run for 30 minutes and 10 seconds, "50s" would run for 50 seconds, etc. Enclose the string in single or double quotes if spaces are used. The MaxRandomInterval is specified in seconds. All alphabet characters will be stripped from the string, so specifying "20s" for a 20 seoond maximum interval is also valid.
    
    
    
    
    -------------------------- EXAMPLE 11 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -StartAt "2024-02-27 11:01:32 AM"
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until February 27th, 2024 at 11:01:32 AM.
    
    
    
    
    -------------------------- EXAMPLE 12 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -StartAt "07/05/2024"
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until July 5th, 2024 at Midnight.
    
    
    
    
    -------------------------- EXAMPLE 13 --------------------------
    
    PS > .\Insomniac-Standalone.ps1 -StartAt "10:30AM"
    
    This example will immediately send a Scroll Lock character, then periodically send a Scroll Lock character at a random interval until stopped with CTRL+C or by closing the program. The random interval changes with each loop. The script will not start sending keypresses until 10:30AM (current day).
    
    
RELATED LINKS
    https://github.com/christopher-conley/insomniac


```

That's all for now.
