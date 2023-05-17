## Intro

<!-- For more icons please follow  https://github.com/MikeCodesDotNET/ColoredBadges -->
<p>
<img src="https://raw.githubusercontent.com/MikeCodesDotNET/ColoredBadges/master/svg/dev/tools/powershell.svg" alt="PowerShell logo"/>

The original, bare-bones version of this was something I put together for a task at work, but after playing with it a bit more, I figured there might be some people out there who find it useful. So here it is.

The use-case is pretty specific and therefore the utility is limited, but you never know when you might need something like this.

### Function and Purpose

1. Takes a SharePoint site URL (e.g. /projects) and combs through the document libraries for directories that contain specified search term in their name
2. Gathers the directory names and short-form URL for each of the above
3. Outputs the collected information and outputs it to a CSV file

### UPDATED Usage

~~In its current form, the script doesn't take arguments~~

✨Parameters now included!✨ You no longer need to alter the script itself to perform the search! Very exciting.

The three required parameters are:

- SiteURL
- SearchTerm
- CSVPath

Example:

`.\sharepoint_search.ps1 -SiteURL "https://contoso.sharepoint.com/sites/MySite" -SearchTerm "MyFolder" -CSVPath "C:\Temp\MyFolder.csv"`

I've also updated the script to include help and usage info directly, which be accessed via `Get-Help`

### Requirements

I'm far from a PS expert, so the script doesn't perform a check for the one required module (PnP.PowerShell) at runtime, though I'd like to figure out how to implement that in the future. I'm still very much in the "grasshopper" stage of using PowerShell for 365 automation/admin.

I _did_ include a .ps1xml file with this requirement listed for documentation purposes.

I run Linux as a base, so everything works with PowerShell 7, but I assume it would work just fine with 5 as well.
