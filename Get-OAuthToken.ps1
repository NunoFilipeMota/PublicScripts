#####################################################################################################
# Function to get OAuth Token using MSAL library
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
# SCRIPT START
$ClientID = "6d32d85f-kdyf-4efb-b325-032a65482549"
$TenantID = "958e9f0e-54ec-3cf2-b5a1-c65215498214"
$CertThumprint = "0F9548B33FC658E324CD5F36854E6E76D325CF3A"

# Retrieve OAuth token
$token = Get-OAuthToken $ClientID $TenantID $CertThumprint

