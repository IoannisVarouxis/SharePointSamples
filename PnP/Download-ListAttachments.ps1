param($SiteUrl, $ListName, $ListItemID = 0, $DownloadPath = ".\")

Import-Module PnP.PowerShell
 
Function Download-ListItemAttachments()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $ListName,
        [Parameter(Mandatory=$true)] [string] $ListItemID,
        [Parameter(Mandatory=$true)] [string] $DownloadPath
    )    
    Try {
 
        #Connect to SharePoint Online site
        Try {
            $ctx = Get-PnPContext
        }
        Catch {
            Connect-PnPOnline -Url $SiteURL -UseWebLogin
        }
        
        #Get the List Item
        $Listitem = Get-PnPListItem -List $ListName -Id $ListItemID
        
        #Get Attachments from List Item
        $Attachments = Get-PnPProperty -ClientObject $Listitem -Property "AttachmentFiles"
         
        #Download All Attachments from List item
        Write-host "Total Number of Attachments Found:"$Attachments.Count
         
        $Attachments | ForEach-Object { 
            Get-PnPFile -Url $_.ServerRelativeUrl -FileName $_.FileName -Path $DownloadPath -AsFile
        }

        Write-Host -f Green "Total List Attachments Downloaded : $($Attachments.count)"
    }
    Catch {
        Write-Host -f Red "Error Downloading List Item Attachments!" $_.Exception.Message
        Write-Host -f Red $_.Exception.StackTrace
    } 
}

 
Function Download-ListAttachments()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $ListName,
        [Parameter(Mandatory=$true)] [string] $DownloadPath,
        [Parameter(Mandatory=$true)] [bool]   $CreateFolderPerItem
    )    
    Try {

        #Connect to SharePoint Online site
        Try {
            $ctx = Get-PnPContext
        }
        Catch {
            Connect-PnPOnline -Url $SiteURL -UseWebLogin
        }
    
        #Get all List Items
        $Listitems = Get-PnPListItem -List $ListName -PageSize 500
         
        #Iterate through List Items
        ForEach($Item in $Listitems)
        {
            #Get Attachments from List Item
            $Attachments = Get-PnPProperty -ClientObject $Item -Property "AttachmentFiles"
         
            #Download All Attachments from List item
            Write-host "Downloading Attachments from List item '$($Item.ID)', Attachments Found: $($Attachments.Count)"
            
            $DownloadLocation = $DownloadPath
            if ($CreateFolderPerItem) {
                #Create directory for each list item
                $DownloadLocation = $DownloadPath+$Item.ID
                If (!(Test-Path -path $DownloadLocation)) { New-Item $DownloadLocation -type Directory | Out-Null }
            }
            $Attachments | ForEach-Object {
                $FileName = $_.FileName
                if (-not $CreateFolderPerItem) {
                    $FileName = $DownloadPath+$Item.id.ToString()+"_"+$_.FileName
                }
                Get-PnPFile -Url $_.ServerRelativeUrl -FileName $FileName -Path $DownloadLocation -AsFile -Force
            }
        }

        Write-Host -f Green "List Attachments Downloaded Successfully!"
    }
    Catch {
        Write-Host -f Red "Error Downloading List Attachments!" $_.Exception.Message
    } 
}

Write-Host "Parameters:"
Write-Host "`tSiteUrl     : $SiteUrl     "
Write-Host "`tListName    : $ListName    "
Write-Host "`tListItemID  : $ListItemID  "
Write-Host "`tDownloadPath: $DownloadPath"

if (-not (Test-Path -Path $DownloadPath)) {
    New-Item -ItemType "directory" -Path $DownloadPath
}

if ($ListItemID -gt 0) {
    Download-ListItemAttachments -SiteURL $SiteUrl -ListName $ListName -ListItemID $ListItemID -DownloadPath $DownloadPath
}
Else {
    Download-ListAttachments -SiteURL $SiteUrl -ListName $ListName -DownloadPath $DownloadPath -CreateFolderPerItem $False
}

