#####################################################################################################
# Function to get OAuth Token using MSAL - Required for uploading files as SDK does not allow that
#####################################################################################################
Function Get-OAuthToken {
    Param ($ClientID, $TenantID, $CertThumprint)

    Try {
        Import-Module MSAL.PS -ErrorAction Stop
    } Catch {
        Write-Host "Unable to import MSAL PowerShell module: '$($_.Exception.Message)'" -ForegroundColor Red
        Exit
    }

    # Get OAuth Token
    Try {
        $ClientCertificate = Get-Item "Cert:\CurrentUser\My\$CertThumprint"
        $token = Get-MsalToken -ClientId $ClientID -TenantId $TenantID -ClientCertificate $ClientCertificate
        Write-Host "Retrieved OAuth token using MSAL"
        Return $token.AccessToken
    } Catch {
        Write-Host "Unable to get OAuth token using MSAL: '$($_.Exception.Message)'" -ForegroundColor Red
        Exit
    }
}


#####################################################################################################
# Function to upload a file to a SharePoint Online library
#####################################################################################################
Function UploadTo-SharePoint {
    Param ($SPsiteID, $SPfolder, $fileName, $file, $token)

    # Files less than 3MB in size can be uploaded directly to SharePoint. If larger, then an upload session is needed
    If ($file.length/1MB -lt 3) {
        $url = "https://graph.microsoft.com/v1.0/sites/$($SPsiteID)/drive/root:/$($SPfolder)/$($fileName):/content"

        Try {
            Invoke-RestMethod -Method PUT -Uri $url -Headers @{Authorization = "Bearer $token"} -InFile $file -ErrorAction Stop | Out-Null
            Write-Host "File uploaded to SharePoint Online: '$fileName' ($([Math]::Round($file.length/1MB, 2))MB)"
        } Catch {
            Write-Host "Unable to upload '$fileName' to SharePoint Online: '$($_.Exception.Message)'" -ForegroundColor Red
            Return $False
        }
    } Else {
        # Construct a POST request to get the upload url
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
        
        # Send bytes to the upload session
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
            Write-Host "File uploaded to SharePoint Online: '$fileName' ($([Math]::Round($file.length/1MB, 0))MB)"
        } Catch {
            Write-Host "Unable to upload '$fileName' to SharePoint Online: '$($_.Exception.Message)'" -ForegroundColor Red
            Return $False
        }
    }
}


#####################################################################################################
# SCRIPT START
$ClientID = "6d32d85f-kdyf-4efb-b325-032a65482549"
$TenantID = "958e9f0e-54ec-3cf2-b5a1-c65215498214"
$CertThumprint = "0F9548B33FC658E324CD5F36854E6E76D325CF3A"

$SPsiteID = "0e5697fc-4ff9-835c-5dc2-65a1a7f65b65"
$SPfolder = "General/Files"

# Retrieve OAuth token, which is requird to upload the file to SharePoint. Unfortunately SDK still doesn't provide a method to do this
$token = Get-OAuthToken $ClientID $TenantID $CertThumprint

# Retrieve the file we want to upload
$file = Get-Item "$PSScriptRoot\TestFile.csv"
UploadTo-SharePoint $SPsiteID $SPfolder "TestFile.csv" $file $token