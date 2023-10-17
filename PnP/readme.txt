
Open PowerShell

Register-PSRepository -Default

Uninstall-Module PnP.PowerShell -AllVersions

# Supports PS version 5.x
Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser
# Supports PS version 7.2+
Install-Module -Name PnP.PowerShell -Scope CurrentUser

    Untrusted repository
    You are installing the modules from an untrusted repository. If you trust this repository, change its
    InstallationPolicy value by running the Set-PSRepository cmdlet. Are you sure you want to install the modules from
    'PSGallery'?
    [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "N"): y


# Call script to get the files for specific item:
.\Download-ListAttachments.ps1 -SiteUrl "https://tenant1.sharepoint.com/sites/site1/" -ListName "List1" -DownloadPath "D:\Temp\SP\Files\"

# Call script to get files for all items:
.\Download-ListAttachments.ps1 -SiteUrl "https://tenant1.sharepoint.com/sites/site1/" -ListName "List1" -ListItemID 10 -DownloadPath "D:\Temp\SP\Files\"

