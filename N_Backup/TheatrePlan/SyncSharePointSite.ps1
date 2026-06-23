#region Functions
function Sync-SharepointLocation {
    param (
        [guid]$siteId,
        [guid]$webId,
        [guid]$listId,
        [mailaddress]$userEmail,
        [string]$webUrl,
        [string]$webTitle,
        [string]$listTitle,
        [string]$syncPath
    )
    try {
        Add-Type -AssemblyName System.Web
        #Encode site, web, list, url & email
        [string]$siteId = [System.Web.HttpUtility]::UrlEncode($siteId)
        [string]$webId = [System.Web.HttpUtility]::UrlEncode($webId)
        [string]$listId = [System.Web.HttpUtility]::UrlEncode($listId)
        [string]$userEmail = [System.Web.HttpUtility]::UrlEncode($userEmail)
        [string]$webUrl = [System.Web.HttpUtility]::UrlEncode($webUrl)
        #build the URI
        $uri = New-Object System.UriBuilder
        $uri.Scheme = "odopen"
        $uri.Host = "sync"
        $uri.Query = "siteId=$siteId&webId=$webId&listId=$listId&userEmail=$userEmail&webUrl=$webUrl&listTitle=$listTitle&webTitle=$webTitle"
        #launch the process from URI
        Write-Host $uri.ToString()
        start-process -filepath $($uri.ToString())
    }
    catch {
        $errorMsg = $_.Exception.Message
    }
    if ($errorMsg) {
        Write-Warning "Sync failed."
        Write-Warning $errorMsg
    } else {
        Write-Host "Sync completed."
        while (!(Get-ChildItem -Path $syncPath -ErrorAction SilentlyContinue)) {
            Start-Sleep -Seconds 2
		}
    return $true
    }    
}
#endregion
#region Main Process
try {
    #region Sharepoint Sync
	$user = [System.Security.Principal.WindowsIdentity]::GetCurrent()
	#Check if AAD account
	$sid = $user.user.value
	[mailaddress]$userUpn = Get-ItemPropertyValue -path "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$sid\IdentityCache\$sid\" -name "UserName"
	Write-Host "User $upn detected"


    $params = @{
        #replace with data captured from your sharepoint site.
        siteId    = "{8f597df3-9810-414e-a5d7-b5d49de73691}"
        webId     = "{7fa0aba5-aa02-4a92-a3c3-f776e58080c2}"
        listId    = "{b18019f8-9bd8-4af9-987a-095039780a9a}"
        userEmail = $userUpn
        webUrl    = "https://theatreplan.sharepoint.com/sites/Templates"
        webTitle  = "Templates"
        listTitle = "Documents"
    }
	$OrganisationDisplayName = "Theatreplan"
	
    $params.syncPath  = "$(split-path $env:onedrive)\" + $OrganisationDisplayName + "\$($params.webTitle) - $($Params.listTitle)"
    Write-Host "SharePoint params:"
    $params | Format-Table
    if (!(Test-Path $($params.syncPath))) {
        Write-Host "Sharepoint folder not found locally, will now sync.." -ForegroundColor Yellow
        $sp = Sync-SharepointLocation @params
        if (!($sp)) {
            Throw "Sharepoint sync failed."
        }
    } else {
        Write-Host "Location already syncronized: $($params.syncPath)" -ForegroundColor Yellow
    }
    #endregion
} catch {
    $errorMsg = $_.Exception.Message
} finally {
    if ($errorMsg) {
        Write-Warning $errorMsg
        Throw $errorMsg
    } else { Write-Host "Completed successfully.."}
}
#endregion

################################################################################################
