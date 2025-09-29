function Get-UpdateStatus{
    
    $needUpdates = "KBArticleID,UpdateId,DownloadUrl`n"

    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
    $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", "$PSScriptRoot\wsusscn2.cab")
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    Write-Host "Searching for updates..."
    $UpdateSearcher.ServerSelection = 3 # ssOthers
    $UpdateSearcher.ServiceID = [string] $UpdateService.ServiceID
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
    $Updates = $SearchResult.Updates
    If ($SearchResult.Updates.Count -eq 0) {
        Write-Host "There are no applicable updates."
        Exit
    }

    Write-Host "List of applicable items on the machine when using wssuscan.cab:"
    For ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
        $update = $SearchResult.Updates.Item($i)
        $downloadUrl = $update.BundledUpdates.item(0).downloadcontents | Where-Object {(($_.DownloadUrl -like ("*$($update.KBArticleIDs)*.msu")) -or ("*$($update.KBArticleIDs)*.exe"))}
        Write-Host ($i + 1) "> " $update.Title
        $needUpdates += "$($update.KBArticleIDs),$($update.Identity.UpdateID),$($downloadUrl.downloadUrl)`n"
    }

    $needUpdates | Out-File -FilePath $PSScriptRoot\needed_updates.csv -Encoding utf8
}


function Invoke-UpdateDownload([string[]]$KBArticleID){
    $requestedUpdates = Import-Csv -Path $PSScriptRoot\needed_updates.csv


    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $Updates =  $UpdateSearcher.search("UpdateID=$($UpdateID)")
    $updateKB = $Updates[0].KBArticleIDs
    $downloadUrl = $Updates[0].BundledUpdates.item(0).downloadcontents | Where-Object {$_.DownloadUrl -like "*$($updateKB)*.msu"}

    if(-not (Test-Path -Path $PSScriptRoot\Downloads\$updateKB)){
        New-Item -Type Directory -Path $PSScriptRoot\Downloads\$updateKB
    }

    Invoke-WebRequest -Uri $downloadUrl -OutFile "$PSScriptRoot\Downloads\$updateKB"
}

