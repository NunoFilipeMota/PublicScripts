#####################################################################################################
# Function to get OAuth Token using MSAL
#####################################################################################################
Function GetOAuthToken {
    Try {
        Import-Module MSAL.PS -ErrorAction Stop
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to import MSAL PowerShell module: '$($_.Exception.Message)'"
        Send-EmailAlert -Body "Unable to import MSAL PowerShell module: '$($_.Exception.Message)'." -Subject "GlobalRelay Automation - ERROR - Unable to import MSAL PowerShell module"
        Return $False
    }

    # Get OAuth Token
    Try {
        $ClientCertificate = Get-Item "Cert:\CurrentUser\My\$CertThumbprint"
        # $ClientCertificate = Get-AutomationCertificate -Name $CertName
        
        $token = Get-MsalToken -ClientId $ClientID -TenantId $TenantID -ClientCertificate $ClientCertificate

        # Get token expiration date and time so we can renew it 2 minutes before it expires
        # $global:tokenExpireDateTime = ((Get-Date).AddSeconds($tokenRequest.expires_in)).AddSeconds(-120)

        # Write-Log -Type "INF" -Message "Retrieved OAuth token."   # Can't run this, otherwise the text will be added to the token variable that called this function
        Return $token.AccessToken
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to get OAuth token using MSAL: '$($_.Exception.Message)'. Exiting script."
        Send-EmailAlert -Body "Unable to get OAuth token using MSAL: '$($_.Exception.Message)'. Exiting script." -Subject "GlobalRelay Automation - ERROR - Unable to get OAuth token using MSAL"
        EndScriptdisc
    }
}


#####################################################################################################
# Function to upload a file to a SharePoint Online library
#####################################################################################################
Function UploadToSharePoint {
    Param ($SPsiteID, $SPfolder, $fileName, $file, $token)

    Write-Host "Uploading '$fileName' ($('{0:N0}' -f $([Math]::Round($file.length/1MB, 0)))MB / $('{0:N0}' -f $([Math]::Round($file.length/1KB, 0)))KB) to SharePoint Online"

    # Files less than 3MB in size can be uploaded directly to SharePoint. If larger, then an upload session is needed
    If ($file.length/1MB -lt 3) {
        $url = "https://graph.microsoft.com/v1.0/sites/$($SPsiteID)/drive/root:/$($SPfolder)/$($fileName):/content"

        Try {
            Invoke-RestMethod -Method PUT -Uri $url -Headers @{Authorization = "Bearer $token"} -InFile $file -ErrorAction Stop | Out-Null
        } Catch {
            Write-Host "Unable to upload '$fileName' to SharePoint Online: '$($_.Exception.Message)'" -ForegroundColor Red
            Return $False
        }
    } Else {
        # Construct a POST request to get the upload URL
        Try {
            $invokeRestMethodParams = @{
                Uri     = "https://graph.microsoft.com/v1.0/sites/$($SPsiteID)/drive/root:/$($SPfolder)/$($fileName):/createUploadSession"
                Method  = "POST"
                Headers = @{
                    Accept         = "application/json"
                    "Content-Type" = "text/plain"
                    Authorization  = "bearer $($token)"
                }
            }

            $response = Invoke-RestMethod @invokeRestMethodParams -ErrorAction Stop
            $uploadUrl = $response.uploadUrl
            If (!$response -or [String]::IsNullOrWhiteSpace($response.uploadUrl)) {
                Write-Host "Unable to retrieve the link to upload the file '$($fileName)' to SharePoint Online" -ForegroundColor Red
                Return $False
            } Else {
                Write-Host "Retrieved the link to upload the file to SharePoint Online"
            }
        } Catch {
            Write-Host "Unable to retrieve the link to upload the file '$($fileName)' to SharePoint Online: '$($_.Exception.Message)'" -ForegroundColor Red
            Return $False
        }
        
        If ($file.length/1MB -lt 249) {
            
            $fileInBytes = [System.IO.File]::ReadAllBytes($file)
            $fileLength = $fileInBytes.Length
            
            $invokeRestMethodParams = @{
                Uri     = $uploadUrl
                Method  = "PUT"
                Body    = $fileInBytes
                Headers = @{'Content-Range' = "bytes 0-$($fileLength-1)/$fileLength"}
            }

            Try {
                $response = Invoke-RestMethod @invokeRestMethodParams -ErrorAction Stop
            } Catch {
                Write-Host "Unable to upload '$fileName' to SharePoint Online: '$($_.Exception.Message)'" -ForegroundColor Red
                Return $False
            }
        } Else {
            #Fragments
            <#
                https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
                To upload the file, or a portion of the file, your app makes a PUT request to the uploadUrl value received 
                in the createUploadSession response. You can upload the entire file, or split the file into multiple byte ranges, 
                as long as the maximum bytes in any given request is less than 60 MiB.

                The fragments of the file must be uploaded sequentially in order. Uploading fragments out of order will result in an error.

                Note: If your app splits a file into multiple byte ranges, the size of each byte range MUST be a multiple of 320 KiB (327,680 bytes). 
                Using a fragment size that does not divide evenly by 320 KiB will result in errors committing some files.
            #>

            $ChunkSize = 62259200
            $reader = [System.IO.File]::OpenRead($file.FullName)
            $buffer = New-Object -TypeName Byte[] -ArgumentList $ChunkSize
            $position = 0
            # Write-Host "ChunkSize: $ChunkSize" -ForegroundColor Cyan
            # Write-Host "BufferSize: $($buffer.Length)" -ForegroundColor Cyan
            $moreData = $True
            While ($moreData) {
                # Read a chunk
                $bytesRead = $reader.Read($buffer, 0, $buffer.Length)
                $output = $buffer
                
                If ($bytesRead -ne $buffer.Length) {
                    # No more data to be read
                    $moreData = $False
                    # Shrink the output array to the number of bytes
                    $output = New-Object -TypeName Byte[] -ArgumentList $bytesRead
                    [Array]::Copy($buffer, $output, $bytesRead)
                    Write-Host "no more data" -ForegroundColor Yellow
                }

                # Upload the chunk
                $Header = @{
                    # 'Content-Length' = $($output.Length)
                    'Content-Range'  = "bytes $position-$($position + $output.Length - 1)/$($file.Length)"
                }

                # Write-Host "Content-Length = $($output.Length)" -ForegroundColor Cyan
                # Write-Host "Content-Range  = bytes $position-$($position + $output.Length - 1)/$($file.Length)" -ForegroundColor Cyan
                $position = $position + $output.Length
                Invoke-RestMethod -Method PUT -Uri $uploadUrl -Body $output -Headers $Header -SkipHeaderValidation #-ContentType "application/octet-stream"
            }
            $reader.Close()
        }
    }

    Write-Host "File successfully uploaded."
}


#####################################################################################################
# SCRIPT START

# Graph API details
$ClientID = "6d32d85f-kdyf-4efb-b264-032a6587452"
$TenantID = "958e9f0e-54ec-l3h5-v6f2-c65215498214"
$CertThumBprint = "0F9548B55GT756K237CD5F36854E6E76D325CF3D"

# SharePoint site details of where to upload files to
$SPsiteID = "rf45h7fe-45s8-40ef-b97e-b61f5e685471"
$SPfolder = "General/Uploads"

# Retrieve OAuth token, which is requird to upload the file to SharePoint. Unfortunately SDK still doesn't provide a method to do this
$token = GetOAuthToken

# Retrieve the file we want to upload
$file = Get-Item "$PSScriptRoot\TestFile.csv"
UploadToSharePoint $SPsiteID $SPfolder "TestFile.csv" $file $token