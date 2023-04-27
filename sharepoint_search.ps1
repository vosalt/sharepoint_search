#Config Variables
$SiteURL = "https://raftelis.sharepoint.com/sites/Projects"
$SearchTerm = "Executive Search"

#Output File
$CSVPath = "./results.csv"

Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive

    #Get All Lists in the Site
    $Lists = Get-PnPList

    #Create an array to hold folder details
    $FolderDetails = @()

    #Loop through all Lists
    foreach ($List in $Lists) {
        #Get All List Items in batches of 500
        $ListItems = Get-PnPListItem -List $List.Title -PageSize 500

        #Filter Folders with Search Term in Name
        $FilteredFolders = $ListItems | 
        Where-Object { $_.FileSystemObjectType -eq "Folder" -and $_["FileLeafRef"] -like "*$SearchTerm*" }
        Write-Host "Number of matching folders in '$($List.Title)': $($FilteredFolders.Count)"

        #Get Folder Details and add to array
        $FolderDetails += $FilteredFolders | ForEach-Object {
            [PSCustomObject]@{
                FolderName = $_.FieldValues.FileLeafRef
                URL        = $_.FieldValues.FileRef
            }
        }
    }

    #Export folder details to CSV
    $FolderDetails | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding utf8 -Delimiter ";"

    Write-Host "Folder details exported to CSV file: $CSVPath"
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}
