<#
.SYNOPSIS
Monitor changes done to an Azure AD group.

.DESCRIPTION
This script monitors any changes done to a specified Azure AD group.
If any changes are detected, it emails those changes to certain users.

.NOTES
This script:
    - Was written to be run in Azure Automation, but can easily be updated to run on a server;
    - Uses Power Automate to send an email alert if Graph API fails. For details on how to configure this, please
      refer to the following article: https://www.linkedin.com/pulse/trigger-power-automate-flow-from-powershell-script-send-nuno-mota

Requirements:
    - Create a variable named "Monitor-AzureADGroupChanges-SearchFrom" as a string with a date in the following format: "MM/dd/yyyy HH:mm:ss"
    - Update with your Azure App Reg details, plus the 'AutomationAccountName' and 'ResourceGroupName' parameters used in the script;
    - PowerShell modules: Microsoft.Graph.Authentication, Microsoft.Graph.Reports, Az.Accounts, and Az.Automation

Author: Nuno Mota

Version: 0.1 - 21/06/2023 - Initial script
#>


#####################################################################################################
# Function to write all the actions performed by the script to a log file
#####################################################################################################
Function Write-Log {
    [CmdletBinding()]
    Param ([String] $Type, [String] $Message)

    Switch ($Type) {
        "INF" {Write-Output $Message}
        "WRN" {Write-Warning $Message}
        "ERR" {Write-Error $Message}
        default {$Message}
    }
}


#####################################################################################################
# Function to connect to Graph API
#####################################################################################################
Function ConnectGraphAPI {
    Try {
        Import-Module Microsoft.Graph.Reports -ErrorAction Stop
        Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -CertificateThumbprint $CertThumprint -ErrorAction Stop
        Write-Log -Type "INF" -Message "Connected to Graph API"
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to connect to Graph API: '$($_.Exception.Message)'"

        # Since we can't connect to Graph API, we send an alert using a Power Automate Flow
        $flowBody = @{
            Subject = "Monitor-AzureADGroupChanges - ERROR - Unable to connect to Graph API"
            Message = "Unable to connect to Graph API: '$($_.Exception.Message)'."
            Secret = $flowSecret
        } | ConvertTo-Json

        Try {
            Invoke-RestMethod -Uri $flowURL -Method POST -Body $flowBody -ContentType "application/json" -ErrorAction Stop
        } Catch {
            Write-Log -Type "ERR" -Message "Unable to trigger Power Automate Flow to send email alert: '$($_.Exception.Message)'."
        }

        EndScript
    }
}


#####################################################################################################
# Function to connect to Azure using Managed Identity
#####################################################################################################
Function ConnectAzure {
    Try {
        Import-Module Az.Accounts -ErrorAction Stop
        Import-Module Az.Automation -ErrorAction Stop

        Connect-AzAccount -Identity -ErrorAction Stop
        Write-Log -Type "INF" -Message "Connected to Azure"
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to connect to Azure: '$($_.Exception.Message)'. Exiting script"
        Send-EmailAlert -Body "Unable to connect to Azure: '$($_.Exception.Message)'. Exiting script" -Subject "Monitor-AzureADGroupChanges - ERROR - Unable to connect to Azure"
        EndScript
    }
}


#####################################################################################################
# Function to send an email (alert)
#####################################################################################################
Function Send-EmailAlert {
    Param ([String] $Body, [String] $Subject)

    $postbody = @{
        Message = @{
            Subject = $Subject
            Body = @{
                ContentType = "HTML"
                Content = $Body
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{Address = "nuno.mota@domain.com"}
                }
                @{
                    EmailAddress = @{Address = "john.doe@domain.com"}
                }
            )
        }
        saveToSentItems = $False
    }

    Try {
        Send-MgUserMail -UserId "nuno.mota@domain.com" -BodyParameter $postbody -ErrorAction Stop
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to send alert email: '$($_.Exception.Message)'"

        # Since we can't send the email alert using Graph API, we send it using a Power Automate Flow
        $flowBody = @{
            Subject = $Subject
            Message = $Body
            Secret = $flowSecret
        } | ConvertTo-Json

        Try {
            Invoke-RestMethod -Uri $flowURL -Method POST -Body $flowBody -ContentType "application/json" -ErrorAction Stop
        } Catch {
            Write-Log -Type "ERR" -Message "Unable to trigger Azure Logic App to send email alert: '$($_.Exception.Message)'."
        }
    }
}


#####################################################################################################
# Function to generate HTML body for group changes email
#####################################################################################################
Function GenerateGroupChangesHTML {
    Param ($events)

    # HTML colours
    $HTMLtableHeader1 = "#002C54" # Dark Blue
    $HTMLtableHeader2 = "#325777" # Lighter Blue
    $HTMLtableHeader3 = "#000000" # Black
    $HTMLred = "#FF0000"
    $HTMLamber = "#FFD000"
    $HTMLgreen = "#00BB00"
    $HTMLfontBlack = "#000000"
    $HTMLfontWhite = "#FFFFFF"
    
    $reportBody =   "<!DOCTYPE html>
                    <HTML>
                    <head>
                    <meta name=""description"" content=""Monitor the changes made to xxxxxxxx group"">
                    <meta name=""author"" content=""Nuno Mota (nuno.mota@domain.com)"">
                    <title>Changes to xxxxxxxx Group Detected</title>

                    <style>
                    body {
                        font-family:Verdana,Arial,sans-serif;
                        font-size: 10pt;
                        background-color: white;
                        color: #000000;
                    }

                    table {
                        border: 0px;
                        border-collapse: separate;
                        padding: 3px;
                    }

                    tr, td, th { padding: 3px; }

                    th {
                        font-weight: bold;
                        text-align: center;
                    }

                    h1,h2,h3,h4,h5 { color: $HTMLtableHeader1; }
                    </style></head>"

    $reportBody +=  "<BODY>
    <h2 align=""center"">Changes to xxxxxxxx Group Detected</h2>
    <h4 align=""center"">$((Get-Date).ToString())</h4>"

    $reportBody += "<br><br><table>
        <tr bgcolor=""$HTMLtableHeader2""; style=""color: white""><th>Date</th><th>Action</th><th>OperationType</th><th>Result</th><th>Requestor</th><th>TargetUser</th></tr>"

    $greyRow = $False
    ForEach ($entry in ($events | Sort Date -Descending)) {
        $reportBody += "<tr"
        If ($greyRow) {$reportBody += " style=""background-color:#dddddd"""; $greyRow = $False} Else {$greyRow = $True}

        $reportBody += "><td>$($entry.Date)</td><td>$($entry.Action)</td><td>$($entry.OperationType)</td><td>$($entry.Result)</td><td>$($entry.Requestor)</td><td>$($entry.TargetUser)</td></tr>"
    }
    $reportBody += "</table><br><br>"

    $reportBody += "<p style=""color:white"">Nuno Mota</p>
    </BODY></HTML>"

    Return $reportBody
}


#####################################################################################################
# Function to terminate the script
#####################################################################################################
Function EndScript {
    Try {Disconnect-MgGraph -ErrorAction Stop} Catch {}
    Try {Disconnect-AzAccount -ErrorAction Stop} Catch {}
    $stopWatch.Stop()
    Write-Log -Type "INF" -Message "END (runtime: $('{0:N0}' -f $($stopWatch.Elapsed.TotalSeconds)) seconds)"
    Exit
}


#####################################################################################################
# START
$stopWatch = [System.Diagnostics.Stopwatch]::startNew()
Write-Log -Type "INF" -Message "START"

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ID of the group we want to monitor
$groupID = "45b143a6 (...)"

# Email Alert Power Automate Flow / Azure Logic App
$flowURL = "https://prod-211.westeurope.logic.azure.com  (...)"
$flowSecret = "YsnMw5 (...)"

# LSEG tenant app registration
$ClientID = "6d00d13f (...)"
$TenantID = "287e9f0e (...)"
$CertThumprint = "0F1046B (...)"

# Connect Graph API and Azure
ConnectGraphAPI
ConnectAzure

# Get the date and time we need to run our search from
Try {
    $searchFrom = (Get-AzAutomationVariable -Name "Monitor-AzureADGroupChanges-SearchFrom" -AutomationAccountName "AZU-(...)" -ResourceGroupName "AZU-(...)" -ErrorAction Stop).Value
    Write-Log -Type "INF" -Message "Searching Audit Logs from  '$searchFrom'"
    Write-Output $searchFrom
    $searchFrom = (Get-Date $searchFrom).ToString("yyyy-MM-ddTHH:mm:ssZ")
} Catch {
    Write-Log -Type "ERR" -Message "Unable to retrieve Azure Automation variable: '$($_.Exception.Message)'. Exiting script"
    Send-EmailAlert -Body "Unable to retrieve Azure Automation variable: '$($_.Exception.Message)'. Exiting script" -Subject "Monitor-AzureADGroupChanges - ERROR - Unable to retrieve Azure Automation variable"
    EndScript
}

# Search Azure AD audit log
Try {
    $logEntries = Get-MgAuditLogDirectoryAudit -All -Filter "LoggedByService eq 'Core Directory' and targetResources/any(t:t/id eq '$groupID') and ActivityDateTime ge $searchFrom" | ? {$_.ActivityDisplayName -ne "Synchronization rule action"}

    # Get current time so we can later update the 'Monitor-AzureADGroupChanges-SearchFrom' variable
    $newSearchFrom = Get-Date -f "MM/dd/yyyy HH:mm:ss"
} Catch {
    Write-Log -Type "ERR" -Message "Unable to query AuditLogDirectoryAudit: '$($_.Exception.Message)'"
    Send-EmailAlert -Body "Unable to query AuditLogDirectoryAudit: '$($_.Exception.Message)'" -Subject "Monitor-AzureADGroupChanges - ERROR - Unable to query AuditLogDirectoryAudit"
}

If ($logEntries) {
    [Array] $events = @()
    ForEach ($event in $logEntries) {
        $events += [PSCustomObject] @{
            Date            = $event.ActivityDateTime
            Action          = $event.ActivityDisplayName
            OperationType   = $event.OperationType
            Result          = $event.Result
            Requestor       = If ($event.InitiatedBy.User.DisplayName) {"$($event.InitiatedBy.User.DisplayName) - $($event.InitiatedBy.User.UserPrincipalName)"} Else {$event.InitiatedBy.User.UserPrincipalName}
            TargetUser      = $event.TargetResources | ? {$_.Id -ne "45b143a6-5179-446b-b5d3-bfff58727ca5"} | Select -ExpandProperty UserPrincipalName
        }
    }

    Write-Log -Type "WRN" -Message "$(($events | Measure).Count) group change(s) detected in Azure AD audit logs"

    # Send email with all the changes found in the logs
    $body = GenerateGroupChangesHTML $events
    Send-EmailAlert -Body $body -Subject "Monitor-AzureADGroupChanges - WARNING - Changes Detected"
} Else {
    Write-Log -Type "INF" -Message "No group changes detected"
}

# Update 'Monitor-AzureADGroupChanges-SearchFrom' variable to current date/time so next time the script runs, it only searches from this point forward
Try {
    Set-AzAutomationVariable -Name "Monitor-AzureADGroupChanges-SearchFrom" -AutomationAccountName "AZU-(...)" -ResourceGroupName "AZU-(...)" -Value $newSearchFrom -Encrypted $False -ErrorAction Stop
    Write-Log -Type "INF" -Message "Variable 'Monitor-AzureADGroupChanges-SearchFrom' updated to '$newSearchFrom'"
} Catch {
    Write-Log -Type "ERR" -Message "Unable to update 'Monitor-AzureADGroupChanges-SearchFrom' variable to '$newSearchFrom': '$($_.Exception.Message)'"
    Send-EmailAlert -Body "Unable to update 'Monitor-AzureADGroupChanges-SearchFrom' variable to '$newSearchFrom': '$($_.Exception.Message)'" -Subject "Monitor-AzureADGroupChanges - ERROR - Unable to update 'Monitor-AzureADGroupChanges-SearchFrom' variable to '$newSearchFrom':"
}

EndScript
