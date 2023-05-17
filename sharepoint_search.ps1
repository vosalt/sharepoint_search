<#
.SYNOPSIS
    Search for folders in all lists on a SharePoint site and output results to CSV file.

.DESCRIPTION
    Search for folders in all lists on a SharePoint site. Site to be searched and search term are entered as parameters. You will be prompted for credentials
    at run time. Ensure you have the PnP PowerShell module installed. The script will output the results to a CSV file containing the name of the folder
    and the current URL.

.PARAMETER SiteURL
    The URL of the SharePoint site to be searched.

.PARAMETER SearchTerm
    The search term to be used to find folders.

.PARAMETER CSVPath
    Output path and filename for generated CSV.

.EXAMPLE
    .\sharepoint_search.ps1 -SiteURL "https://contoso.sharepoint.com/sites/MySite" -SearchTerm "MyFolder" -CSVPath "C:\Temp\MyFolder.csv"


#>

# Parameters for configuring site, search, and output variables
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,

    [Parameter(Mandatory = $true)]
    [string]$SearchTerm,

    [Parameter(Mandatory = $true)]
    [string]$CSVPath
)

Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive

    # Gather all lists on the specified site
    $Lists = Get-PnPList

    # Create an array to hold folder details
    $FolderDetails = @()

    # Loop through all Lists
    foreach ($List in $Lists) {
        # Get all list items in batches of 500 to avoid list view threshold
        $ListItems = Get-PnPListItem -List $List.Title -PageSize 500

        # Filter folders with search term in name
        $FilteredFolders = $ListItems | 
        Where-Object { $_.FileSystemObjectType -eq "Folder" -and $_["FileLeafRef"] -like "*$SearchTerm*" }
        Write-Host "Number of matching folders in '$($List.Title)': $($FilteredFolders.Count)"

        # Get the desired details (folder name and URL) and add to array
        $FolderDetails += $FilteredFolders | ForEach-Object {
            [PSCustomObject]@{
                FolderName = $_.FieldValues.FileLeafRef
                URL        = $_.FieldValues.FileRef
            }
        }
    }

    # Export gathered details to CSV; using semicolon as delimiter to ensure Excel compatibility
    $FolderDetails | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding utf8 -Delimiter ";"

    Write-Host "Folder details exported to CSV file: $CSVPath"
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}
