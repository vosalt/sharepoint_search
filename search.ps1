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

    #Loop through all Lists
    foreach ($List in $Lists) {
        #Write-Host "Searching list: $($List.Title)"
        #Get All List Items in batches of 500
        $ListItems = Get-PnPListItem -List $List.Title -PageSize 500

        #Filter Folders with Search Term in Name
        $FilteredFolders = $ListItems | Where-Object { $_.FileSystemObjectType -eq "Folder" -and $_["FileLeafRef"] -like "*$SearchTerm*" }
        Write-Host "Number of matching folders: $($FilteredFolders.Count)"

        #Get Folder Details and Export to CSV
        $FolderDetails = $FilteredFolders | Select-Object Title, FileRef
        $FolderDetails | Export-Csv -Path $CSVPath -NoTypeInformation -Append -Encoding utf8
    }

    Write-Host "Folder details exported to CSV file: $CSVPath"
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}