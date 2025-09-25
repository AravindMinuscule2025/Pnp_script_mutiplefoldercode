Import-Module PnP.PowerShell

# Access token Variables
$siteId = ""
$tenantId = ""
$clientId = ""
$clientSecret = ""

$global:expiryTime = $null
#Access token function 
function Get-AccessToken {
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $clientId
        client_secret = $clientSecret
    }


    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body
    $token = $tokenResponse.access_token
    $expiresIn = $tokenResponse.expires_in   #  3599 (1 hour)
    $global:expiryTime = (Get-Date).AddSeconds($expiresIn)
    Write-Host "Access token acquired. Expires at $global:expiryTime" -ForegroundColor Green
    return $token


}
#token generation
$token = Get-AccessToken

# Array to collect error messages
$errorMessages = @()
# Function to send error email
function Send-ErrorMail {
    param(
        [string]$fromUser,
        [string]$toUser,
        [string]$subject,
        [array]$messages,
        [string]$token
    )

    $htmlBody = @"
<html>
   <body>
   Hi Team,<br><br>
    Please find below the errors identified during the file move process
    <h2 style='color:red;'> Errors Summary:</h2>
        <ol>
"@
    # Group messages by Category so category is printed once
    $grouped = $messages | Group-Object Category

    foreach ($g in $grouped) {
        $htmlBody += "<h3 style='color:blue;'>$($g.Name) ($($g.Count))</h3><ol>"

        foreach ($msg in $g.Group) {
            # Replace newlines in the Message property only
            $msgHtml = $msg.Message -replace "`n", "<br>"
            $htmlBody += "<li>$msgHtml</li>"
        }

        $htmlBody += "</ol><br>"
    }

    # End HTML
    $htmlBody += @"
    <p>Please review and take necessary action.</p>
  </body>
</html>
"@

    $mail = @{
        message         = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $htmlBody
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $toUser } }
            )
        }
        saveToSentItems = "true"
    } | ConvertTo-Json -Depth 5

    $uri = "https://graph.microsoft.com/v1.0/users/$fromUser/sendMail"

    try {
        Invoke-WebRequest -Method Post -Uri $uri `
            -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
            -Body $mail | Out-Null
        Write-Host "Email sent successfully: $subject" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
    }
}



# 1. Get items from source list
$listName = "Europe_source_folderbasedfile_list"
$listItemsApi = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listName/items?expand=fields&top=5000"
$listItemsAll = Invoke-RestMethod -Uri $listItemsApi -Headers @{ Authorization = "Bearer $token" }
$listItems = $listItemsAll.value

Write-Host $listItems.Count

# 2. Get files from source library
$libraryDriveId = "b!BzAxdNAlSUOpTE-0cKY5cx_xrr1o_f1FhUdZHct2g2ua7Fx5IRTDS7QelY_3lkpm"
$filesApi = "https://graph.microsoft.com/v1.0/drives/$libraryDriveId/root/children?top=5000"
$files = Invoke-RestMethod -Uri $filesApi -Headers @{ Authorization = "Bearer $token" }

# Build hashtable for file names
$targetFileNames = @{}
foreach ($file in $files.value) {
    $fileName = $file.Name
    $fileId = $file.id      
    $fileSize = $file.size      # File size in bytes
    $targetFileNames[$fileName] = [PSCustomObject]@{
        Id   = $fileId
        Size = $fileSize
    }
}

# 3. Get overall site collection folder and site url list 
$folderRefList = "foldercode&siteurl"
$folderRefItemsApi = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$folderRefList/items?expand=fields&top=5000"
$folderRefItems = Invoke-RestMethod -Uri $folderRefItemsApi -Headers @{ Authorization = "Bearer $token" }

# Hashtable for foldercode -> siteurl
$folderCodeToUrl = @{}
foreach ($ref in $folderRefItems.value) {
    $code = $ref.fields.field_3
    $url = $ref.fields.field_2
    $folderCodeToUrl[$code] = $url
   
}

# 4. Get all site document libraries and their IDs
$allSiteItems = @()
$AllSiteDetailsList = "SiteList&Library-ID'S"
$siteListItemsApi = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$AllSiteDetailsList/items?expand=fields&top=1000"
$siteListItemsAll = Invoke-RestMethod -Uri $siteListItemsApi -Headers @{ Authorization = "Bearer $token" }
foreach ($item in $siteListItemsAll.value) {
    $obj = [PSCustomObject]@{
        Title    = $item.fields.Title
        ListID   = $item.fields.field_2
        siteID   = $item.fields.field_4
        SiteName = $item.fields.field_5
         
    }

    # Add to array
    $allSiteItems += $obj
}

# 5. Process each item in source list

$itemsWithSite = @()

foreach ($item in $listItems) {
    $fields = $item.fields

    $title = $fields.Title
    $name1 = $fields.Name1
    $dueDate = $fields.DueDate
    $status = $fields.Status
    $folderCodes = $fields.folder_code

    # Take only the first code (ignore others)
    $folderName = ($folderCodes -split ",")[0].Trim()

    # --- Validate FolderCode ---
    if (-not $folderCodeToUrl.ContainsKey($folderName)) {
        $errorMessages += [PSCustomObject]@{
            Category = "Validation Error"
            Message  = "<b>FolderCode not found in reference list</b>`n  FileName: $title`n FolderCode: $folderName"
        }

        continue
    }

    # --- Validate Title exists in target file names ---
    if (-not $targetFileNames.ContainsKey($title)) {
        $errorMessages += [PSCustomObject]@{
            Category = "Validation Error"
            Message  = "<b>File not found in targetFileNames list</b>`n  FileName: $title`n "
        }
        
        continue
    }
    # --- Validate Name1 is not null/empty ---

    if ([string]::IsNullOrWhiteSpace($name1)) {
        $errorMessages += [PSCustomObject]@{
            Category = "Validation Error"
            Message  = "<b>Validation failed: Name1 field is empty</b>`n  FileName: $title`n  FolderCode: $folderName"
        }
        

        continue
    }    
    # --- Validate Completed status with multiple folder codes ---
    if ($Status -eq "Completed" -and $foldercodes.Count -ne 1) {
        $errorMessages += [PSCustomObject]@{
            Category = "Validation Error"
            Message  = "<b>Validation failed: Completed item must have exactly 1 FolderCode</b>`n  FileName: $title`n  FolderCode: $folderCodes"
        }
        continue
    }

    # If all validations pass, process the item
    $itemsWithSite += [PSCustomObject]@{
        Title      = $title
        Name1      = $name1
        DueDate    = $dueDate
        Status     = $status
        FolderCode = $folderCodes
        SiteUrl    = $folderCodeToUrl[$folderName]
    }
    
}
Write-Host $itemsWithSite.Count

# Group items by SiteUrl
$listItemsnew = $itemsWithSite | Group-Object SiteUrl

$itemStartTime = Get-Date

Write-Host ">>> Started processing '$Title' at $itemStartTime" -ForegroundColor Gray


foreach ($siteGroup in $listItemsnew) {
    $siteUrl = $siteGroup.Name
    $groupItems = $siteGroup.Group  # All items in this site group
     

    foreach ($item in $groupItems) {
        
        write-Host "Processing item: $($item.count) items for site $siteUrl" -ForegroundColor Magenta
        #token generation
        $now = Get-Date
        if ($now -ge $global:expiryTime.AddMinutes(-5)) {
            $token = Get-AccessToken
        }
        else {
            Write-Host "Using valid token. Expires at $global:expiryTime" -ForegroundColor Yellow
        }

        $Title = $item.Title.Trim()
        $codes = $item.FolderCode.Trim()
        $Name1 = $item.Name1
        $DueDate = $item.DueDate
        $Status = $item.Status
        Write-Host "Now processing $Title ($Status) for $siteUrl ($codes)" -ForegroundColor Cyan

        # Split folder_code by comma and trim
        $code_splitbycomma = $codes -split ',' | ForEach-Object { $_.Trim() }
        $foldercodes = $code_splitbycomma

        $matchSourceDrive = $allSiteItems | Where-Object { $_.Title -eq "Europe_source_files" }
        if ($null -eq $matchSourceDrive) {
            Write-Host "Source drive 'Europe_source_files' not found. Skipping site $siteUrl" -ForegroundColor Red
            continue
        }

        $sourceDriveId = $matchSourceDrive.ListID

        # Get target drive ID for "Documents" library in this site
        $matchTargetDrive = $allSiteItems | Where-Object { $_.SiteName -eq $siteUrl -and $_.Title -eq "Documents" }
 
        if ($null -eq $matchTargetDrive) {
            Write-Host "Target drive 'Documents' not found for site $siteUrl. Skipping." -ForegroundColor Red
            continue
        }

        $targetDriveId = $matchTargetDrive.ListID

        $allSucceeded = $true

        foreach ($foldercode in $foldercodes ) {

            if (-not $folderCodeToUrl.ContainsKey($foldercode)) {
                Write-Host " Folder code '$foldercode' not found in reference list." -ForegroundColor Yellow
                continue
            }
            
            try {
            
                $foldername = "property-$foldercode"
              
                $fileInfo = $targetFileNames[$Title]
                if (-not $fileInfo) {
                    Write-Host "No file info found for '$Title' — skipping." -ForegroundColor DarkYellow
                    continue
                }
                $sourceItemId = $fileInfo.Id
                $fileSize = $fileInfo.Size   
           
                # Declare globally before if/else
                $response = $null
                $fileId = $null
                
   

               
                        
                #less then 50 mb put
                if ($fileSize -le 50MB) {
                    $downloadUrl = "https://graph.microsoft.com/v1.0/drives/$sourceDriveId/items/$sourceItemId/content"
                    $fileBytes = Invoke-WebRequest -Uri $downloadUrl -Headers @{ Authorization = "Bearer $token" } -Method GET

                    # === 4. Upload to target folder ===
                    $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$targetDriveId/root:/$foldername/${Title}:/content"
                    $maxUploadTries = 5
 
                    for ($uploadAttempt = 1; $uploadAttempt -le $maxUploadTries; $uploadAttempt++) {
                        try {
                            Write-Host "Attempt $uploadAttempt : Uploading '$Title' ($fileSize MB)..." -ForegroundColor Yellow

                            $response = Invoke-WebRequest -Uri $uploadUrl `
                                -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/octet-stream" } `
                                -Method PUT -Body $fileBytes.Content
                            if ($response.StatusCode -eq 201 -or $response.StatusCode -eq 200) {
                                Write-Host "Upload successful (StatusCode: $($response.StatusCode))" -ForegroundColor Green
                                $responseJson = $response.Content | ConvertFrom-Json
                                $fileId = $responseJson.id
                                break  #  exit loop once successful
                            }
                           
                        }
                        catch {
                            $errordetails = $_.Exception.Message 
                            Write-Host "Upload failed on attempt $uploadAttempt → $_" -ForegroundColor Red
                            if ($uploadAttempt -lt $maxUploadTries) {
                                Write-Host "Retrying in 1 seconds..." -ForegroundColor DarkYellow
                                Start-Sleep -Seconds 1
                            }
                            else {
                                throw  $errordetails 
                            }
                        }
                    }
                }
                else {
                
                    Write-Host "Copying large file ($fileSize bytes) → using Graph copy API"

                    $moveApi = "https://graph.microsoft.com/v1.0/drives/$sourceDriveId/items/$sourceItemId/copy?@microsoft.graph.conflictBehavior=replace"

                    $body = @{
                        parentReference = @{
                            driveId = $targetDriveId
                            path    = $foldername
                        }

                        name            = $Title

                    } | ConvertTo-Json -Depth 3
                 
                    $headers = @{
                        "Authorization" = "Bearer $token"
                        "Content-Type"  = "application/json"
                        "Prefer"        = "respond-async"
                    }
                    $maxCopyTries = 5
                    for ($copyAttempt = 1; $copyAttempt -le $maxCopyTries; $copyAttempt++) {
                        try {
                            Write-Host "Attempt $copyAttempt : Sending Copy API request..." -ForegroundColor Yellow

                            $response = Invoke-WebRequest -Uri $moveApi -Method POST -Headers $headers -Body $body
                            if ($response.StatusCode -eq 202 -or $response.StatusCode -eq 200) {
                                Write-Host "Copy request accepted by Graph (StatusCode: $($response.StatusCode))" -ForegroundColor Green
                                break  # Exit loop once successful
                            }
                        
                        }
                        catch {
                            $errordetails = $_.Exception.Message
                            Write-Host "Copy API request failed on attempt $copyAttempt → $_" -ForegroundColor Red
                            if ($copyAttempt -lt $maxCopyTries) {
                                Write-Host "Retrying in 5 seconds..." -ForegroundColor DarkYellow
                                Start-Sleep -Seconds 5
                            }
                            else {
                                throw $errordetails
                            }
                        }
                            
                    }
                    # Large file, async copy → poll until completed
                    $locationUrl = $response.Headers.Location[0]
                    $completed = $false
                    $maxTries = 5
                    $attempt = 0
                    for ($attempt = 1; $attempt -le $maxTries; $attempt++) {
                        $sizeMB = [math]::Round($fileSize / 1MB, 2)

                        $pollInterval = [math]::Ceiling(($sizeMB / 50) * 5 * 4) # Poll interval in seconds
                    
                        $startTime = Get-Date  # Track attempt start
                        while (-not $completed) {
                            $now = Get-Date
                            $elapsed = ($now - $startTime).TotalSeconds
                        
                            # Exit this attempt if it exceeds pollInterval threshold
                            if ($elapsed -ge $pollInterval) {
                                Write-Host "Exceeded $pollInterval sec in attempt $attempt, moving to next..." -ForegroundColor DarkYellow
                                break
                            }
                        
                            try {
                                Start-Sleep -Seconds 5
                                Write-Host "Sleeping $pollInterval sec before next poll (FileSize: $sizeMB MB)" -ForegroundColor Gray
                                $fileStatus = Invoke-WebRequest -Uri $locationUrl
                                $statusJson = $fileStatus.Content | ConvertFrom-Json
                                Write-Host "Copy progress for '$Title': $($statusJson.percentComplete)%" -ForegroundColor Cyan
                                if ($statusJson.status -eq "completed") {
                                    Write-Host " Copy Task completed for '$Title'" -ForegroundColor Green
                                    $completed = $true
                                }
                                elseif ($statusJson.status -eq "failed") {
                                    throw " Copy failed for $Title → $($statusJson.error.message)"
                                }
                            }
                            catch {
                                $errordetails = $_.Exception.Message
                                break  # Exit while loop to retry
                            }
                        }
                
                    }
                    if (-not $completed) {
                        throw " Copy did not complete after $maxTries attempts for '$Title' → $errordetails"
                    }
            
                    # Now lookup file in target drive
                    $searchUrl = "https://graph.microsoft.com/v1.0/drives/$targetDriveId/root:/$foldername/$Title"
                    $fileResponse = Invoke-RestMethod -Uri $searchUrl -Headers @{ Authorization = "Bearer $token" } -Method GET
                    $fileId = $fileResponse.id
                   
                    
                }
                if ($null -ne $fileId) {
                    $metadata = @{
                        Name1   = $Name1
                        DueDate = $DueDate
                        Status  = $Status
                    } | ConvertTo-Json -Depth 3

                    Write-Host "Metadata to update: $metadata" -ForegroundColor Gray

                    $updateUrl = "https://graph.microsoft.com/v1.0/drives/$targetDriveId/items/$fileId/listItem/fields"

                    $maxRetries = 5
                    $updated = $false

                    for ($attempt = 1; $attempt -le $maxRetries -and -not $updated; $attempt++) {
                        try {
                            Invoke-RestMethod -Uri $updateUrl `
                                -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
                                -Method PATCH -Body $metadata

                            Write-Host " Metadata updated for '$Title' (Attempt $attempt)" -ForegroundColor Green
                            $updated = $true
                        }
                        catch {
                            $errordetails = $_.ErrorDetails.Message 


                            if ($attempt -lt $maxRetries) {
                                Start-Sleep -Seconds 1  # wait 1 seconds before retrying
                            }
                            else {
                                throw $errordetails
                            }
                        }
                    }

                   
                }


              
                
                    
            }
            catch {
              
                $errorMessage = $_.Exception.Message

                $errorMessages += [PSCustomObject]@{
                    Category = "Processing Error"
                    Message  = "<b>Processing Error</b>`n  FileName: $title`n FolderCode: $folderName `n  ErrorMessage: <span style='color:red'>$errorMessage</span>"
                }
                $allSucceeded = $false
                continue
            }
            

            

            
            if ($allSucceeded) {
            

                # Example log update
                $logApi = "https://graph.microsoft.com/v1.0/sites/root:/sites/Europe:/lists/ProcessedfilesLogItems/items"
                $logBody = @{
                    fields = @{
                        Title      = $Title
                        foldercode = $codes
                        Name1      = $Name1
                        DueDate    = $DueDate
                        Status     = $Status
                    }
                } | ConvertTo-Json -Depth 3
                $maxLogTries = 5

                for ($logAttempt = 1; $logAttempt -le $maxLogTries; $logAttempt++) {
                    try {
                        Write-Host "Attempt $logAttempt : Writing log entry for '$Title'..." -ForegroundColor Yellow


                        Invoke-RestMethod -Uri $logApi `
                            -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
                            -Method POST -Body $logBody
                        Write-Host "Log entry successfully created for '$Title'" -ForegroundColor Green
                        break
                    }
                    catch {
                        $errordetails = $_.Exception.Message
                        $errorMessages += [PSCustomObject]@{
                            Category = "Processing Error"
                            Message  = "<b>Processing Error</b>`n  FileName: $title`n FolderCode: $folderName `n    ErrorMessage: <span style='color:red'>$errordetails</span>"
                        }
                        if ($logAttempt -lt $maxLogTries) {
                            Write-Host "Retrying in 5 seconds..." -ForegroundColor DarkYellow
                            Start-Sleep -Seconds 5
                        }
                        else {
                            Write-Host "Log update failed after $maxLogTries attempts for '$Title'" -ForegroundColor Red
                        }
                    }
                }
            }
      



        
        }

    }
}

if ($errorMessages.Count -gt 0) {
    $now = Get-Date
    if ($now -ge $global:expiryTime.AddMinutes(-5)) {
        $token = Get-AccessToken
    }
    else {
        Write-Host "Using valid token. Expires at $expiryTime" -ForegroundColor Yellow
    }

    Send-ErrorMail -fromUser "e207@minusculetechnologies.com" `
        -toUser "e207@minusculetechnologies.com" `
        -subject "File Move Script (Content & Copy API) -  Errors ($($errorMessages.Count) issues)" `
        -messages $errorMessages `
        -token $token
}

$itemEndTime = Get-Date
$itemDuration = New-TimeSpan -Start $itemStartTime -End $itemEndTime
Write-Host "<<< Finished processing '$Title' at $itemEndTime (Duration: $($itemDuration.TotalSeconds) seconds)" -ForegroundColor Gray

    





