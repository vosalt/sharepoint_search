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

### Usage

In its current form, the script doesn't take arguments (I'm considering updating it, but PowerShell isn't my specialty), so you'll have to edit the script itself to include:

- Site URL
- Search term
- Output file path

### Requirements

There's an included .ps1xml file that lists the one and only module required (PnP.PowerShell) for this to work. As mentioned, PowerShell isn't my specialty so the script won't install/import this module automatically. I included the .ps1xml for documentation purposes, as well as for anyone who knows how to properly implement them in the code.

I am but a humble newcomer to PowerShell.

I run Linux as a base, so everything works with PowerShell 7, but I assume it would work just fine with 5 as well.
