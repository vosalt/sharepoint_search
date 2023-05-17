# Configure site and search variables
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,

    [Parameter(Mandatory = $true)]
    [string]$SearchTerm
)

# Set output file path
$CSVPath = "<your-file-path.csv>"

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
