﻿<#
.SYNOPSIS
Monitor changes made to PIM roles.

.DESCRIPTION
This scripts searches the Azure AD logs for changes made to certain PIM roles. If any are found,
it emails those changes.

.NOTES
This script:
    1. Is designed to be run in Azure Automation;
    2. Uses Graph API to search Azure AD logs, so an app registration with 'AuditLog.Read.All' or 'Directory.Read.All' permissions;
    3. Requires the Graph API PowerShell SDK modules 'Microsoft.Graph.Reports' and 'Microsoft.Graph.Users';
    4. Uses Power Automate to send an email alert if Graph API fails. For details on how to configure this, please
       refer to the following article: https://www.linkedin.com/pulse/trigger-power-automate-flow-from-powershell-script-send-nuno-mota
    5. Requires an Azure Automation Variable named 'Monitor-PIMRoles-SearchFrom' configured as a string with a date in the following format: "MM/dd/yyyy HH:mm:ss";
    6. Update with your Azure App Reg details, plus the 'AutomationAccountName' and 'ResourceGroupName' parameters used in the script;

Author:     Nuno Mota
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
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -CertificateThumbprint $CertThumprint -ErrorAction Stop
        Write-Log -Type "INF" -Message "Connected to Graph API"
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to connect to Graph API: '$($_.Exception.Message)'"

        # Since we can't connect to Graph API, we send an alert using a Power Automate Flow
        $flowBody = @{
            Subject = "Monitor-PIMRoles - ERROR - Unable to connect to Graph API"
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
        Send-EmailAlert -Body "Unable to connect to Azure: '$($_.Exception.Message)'. Exiting script" -Subject "Monitor-CloudShellGroup - ERROR - Unable to connect to Azure"
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
                    EmailAddress = @{Address = "nuno@domain.com"}
                }
            )
        }
        saveToSentItems = $False
    }

    Try {
        Send-MgUserMail -UserId "nuno@domain.com" -BodyParameter $postbody -ErrorAction Stop
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to send alert email: '$($_.Exception.Message)'"

        # Since we can't send the alert using Graph API, we send it using a Power Automate Flow instead
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
Function GenerateRoleChangesHTML {
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
                    <meta name=""description"" content=""Monitor the changes made to key PIM Roles"">
                    <meta name=""author"" content=""Nuno Mota (nuno@domain.com)"">
                    <title>Changes to PIM Roles Detected</title>

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
    <h2 align=""center"">Changes to PIM Roles Detected</h2>
    <h4 align=""center"">$((Get-Date).ToString())</h4>"

    $reportBody += "<br><br><table>
        <tr bgcolor=""$HTMLtableHeader2""; style=""color: white""><th>Date</th><th>Action</th><th>Result</th><th>Requestor</th><th>Role</th><th>Target</th><th>Reason</th></tr>"

    $greyRow = $False
    ForEach ($entry in $events | Sort Date) {
        $reportBody += "<tr"
        If ($greyRow) {$reportBody += " style=""background-color:#dddddd"""; $greyRow = $False} Else {$greyRow = $True}

        $reportBody += "><td>$($entry.Date)</td><td>$($entry.Action)</td><td>$($entry.Result)</td><td>$($entry.Requestor)</td><td>$($entry.Role)</td><td>$($entry.Target)</td><td>$($entry.Reason)</td></tr>"
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
# Script Start
#####################################################################################################
$stopWatch = [System.Diagnostics.Stopwatch]::startNew()
Write-Log -Type "INF" -Message "START"

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Email Alert Power Automate Flow / Azure Logic App
$flowURL = "https://prod-211.westeurope.logic.azure.com:443/workflows/e35c82a7(...)"
$flowSecret = "YsnMw5(...)"

# Azure App Reg details
$ClientID = "6d00d13f-(...)"
$TenantID = "287e9f0e-(...)"
$CertThumprint = "0F1046(...)"

# PIM Roles to monitor (Get-AzureADMSPrivilegedRoleDefinition -ProviderId aadRoles -ResourceId $TenantID | Sort DisplayName | FT DisplayName, ID)
$RoleID = @{}
$RoleID["Authentication Administrator"]             = "c4e39bd9-1100-46d3-8c65-fb160da0071f"
$RoleID["Cloud App Security Administrator"]         = "892c5842-a9a6-463a-8041-72aa08ca3cf6"
$RoleID["Compliance Administrator"]                 = "17315797-102d-40b4-93e0-432062caca18"
$RoleID["Compliance Data Administrator"]            = "e6d1a23a-da11-4be4-9570-befc86d067a7"
$RoleID["Conditional Access Administrator"]         = "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9"
$RoleID["Directory Readers"]                        = "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
$RoleID["Exchange Administrator"]                   = "29232cdf-9323-42fd-ade2-1d097af3e4de"
$RoleID["Exchange Recipient Administrator"]         = "31392ffb-586c-42d1-9346-e59415a2cc4e"
$RoleID["Fabric Administrator"]                     = "a9ea8996-122f-4c74-9520-8edcd192826c"
$RoleID["Global Administrator"]                     = "62e90394-69f5-4237-9190-012177145e10"
$RoleID["Global Reader"]                            = "f2ef992c-3afb-46b9-b7cf-a126ee74c451"
$RoleID["Groups Administrator"]                     = "fdd7a751-b60b-444a-984c-02652fe8fa1c"
$RoleID["Helpdesk Administrator"]                   = "729827e3-9c14-49f7-bb1b-9608f156bbb8"
$RoleID["Insights Administrator"]                   = "eb1f4a8d-243a-41f0-9fbd-c7cdf6c5ef7c"
$RoleID["Insights Analyst"]                         = "25df335f-86eb-4119-b717-0ff02de207e9"
$RoleID["Insights Business Leader"]                 = "31e939ad-9672-4796-9c2e-873181342d2d"
$RoleID["Knowledge Administrator"]                  = "b5a8dcf3-09d5-43a9-a639-8e29ef291470"
$RoleID["Knowledge Manager"]                        = "744ec460-397e-42ad-a462-8b3f9747a02c"
$RoleID["License Administrator"]                    = "4d6ac14f-3453-41d0-bef9-a3e0c569773a"
$RoleID["Message Center Privacy Reader"]            = "ac16e43d-7b2d-40e0-ac05-243ff356ab5b"
$RoleID["Message Center Reader"]                    = "790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b"
$RoleID["Office Apps Administrator"]                = "2b745bdf-0803-4d80-aa65-822c4493daac"
$RoleID["Power Platform Administrator"]             = "11648597-926c-4cf3-9c36-bcebb0ba8dcc"
$RoleID["Privileged Authentication Administrator"]  = "7be44c8a-adaf-4e2a-84d6-ab2649e08a13"
$RoleID["Privileged Role Administrator"]            = "e8611ab8-c189-46e8-94e1-60213ab1f814"
$RoleID["Reports Reader"]                           = "4a5d8f65-41da-4de4-8968-e035b65339cf"
$RoleID["Security Administrator"]                   = "194ae4cb-b126-40b2-bd5b-6091b380977d"
$RoleID["Security Operator"]                        = "5f2222b1-57c3-48ba-8ad5-d4759f1fde6f"
$RoleID["Security Reader"]                          = "5d6b6bb7-de71-4623-b4af-96380a352509"
$RoleID["Service Support Administrator"]            = "f023fd81-a637-4b56-95fd-791ac0226033"
$RoleID["SharePoint Administrator"]                 = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
$RoleID["Skype for Business Administrator"]         = "75941009-915a-4869-abe7-691bff18279e"
$RoleID["Teams Administrator"]                      = "69091246-20e8-4a56-aa4d-066075b2a7a8"
$RoleID["Teams Communications Administrator"]       = "baf37b3a-610e-45da-9e62-d9d1e5e8914b"
$RoleID["Teams Communications Support Engineer"]    = "f70938a0-fc10-4177-9e90-2178f8765737"
$RoleID["Teams Communications Support Specialist"]  = "fcf91098-03e3-41a9-b5ba-6f0ec8188a12"
$RoleID["Usage Summary Reports Reader"]             = "75934031-6c7e-415a-99d7-48dbd49e875e"
$RoleID["User Administrator"]                       = "fe930be7-5e62-47db-91af-98c3a49a38b1"
$RoleID["Viva Goals Administrator"]                 = "92b086b3-e367-4ef2-b869-1de128fb986e"
$RoleID["Viva Pulse Administrator"]                 = "87761b17-1ed2-4af3-9acd-92a150038160"
$RoleID["Yammer Administrator"]                     = "810a2642-a034-447f-a5e8-41beaa378541"


# Connect Graph API and Azure
ConnectGraphAPI
ConnectAzure


# Check for PIM role assignment changes
$ActivityDisplayName = @("Add member to role in PIM completed (permanent)", "Add eligible member to role in PIM completed (permanent)", "Add member to role outside of PIM (permanent)", "Remove eligible member from role in PIM completed (permanent)", "Remove member from role in PIM completed (permanent)", "Remove permanent eligible role assignment", "Remove permanent direct role assignment")

# Get the date and time we need to run our search from
Try {
    $searchFrom = (Get-AzAutomationVariable -Name "Monitor-PIMRoles-SearchFrom" -AutomationAccountName "AZU-(...)" -ResourceGroupName "AZU-(...)" -ErrorAction Stop).Value
    # $searchFrom = (Get-Date).AddDays(-15)
    Write-Log -Type "INF" -Message "Searching Audit Logs from '$searchFrom'"
    $searchFrom = (Get-Date $searchFrom).ToString("yyyy-MM-ddTHH:mm:ssZ")
} Catch {
    Write-Log -Type "ERR" -Message "Unable to retrieve Azure Automation variable: '$($_.Exception.Message)'. Exiting script"
    Send-EmailAlert -Body "Unable to retrieve Azure Automation variable: '$($_.Exception.Message)'. Exiting script" -Subject "Monitor-PIMRoles - ERROR - Unable to retrieve Azure Automation variable"
    EndScript
}

# Retrieve PIM logs from Azure AD audit log
Try {
    $logEntries = Get-MgAuditLogDirectoryAudit -All -Filter "LoggedByService eq 'PIM' and ActivityDateTime ge $searchFrom" -ErrorAction Stop | ? {$ActivityDisplayName -contains $_.ActivityDisplayName}

    # Get current time so we can later update the 'Monitor-CloudShellGroup-SearchFrom' variable
    $newSearchFrom = Get-Date -f "MM/dd/yyyy HH:mm:ss"
} Catch {
    Write-Log -Type "ERR" -Message "Unable to query AuditLogDirectoryAudit: '$($_.Exception.Message)'"
    Send-EmailAlert -Body "Unable to query AuditLogDirectoryAudit: '$($_.Exception.Message)'" -Subject "Monitor-PIMRoles - ERROR - Unable to query AuditLogDirectoryAudit"
    EndScript
}

If ($logEntries) {
    [Array] $PIMevents = @()
    ForEach ($event in $logEntries) {
        $PIMevents += [PSCustomObject] @{
            Date            = $event.ActivityDateTime
            Action          = $event.ActivityDisplayName
            # OperationType   = $event.OperationType
            Result          = $event.Result
            # Reason        = $event.AdditionalDetails.Justification
            Reason          = $event.ResultReason
            # TicketNumber    = $event.AdditionalDetails.TicketNumber
            Requestor       = If ($event.InitiatedBy.User.UserPrincipalName) {$event.InitiatedBy.User.UserPrincipalName} Else {$event.InitiatedBy.App.DisplayName}
            Role            = ($event.TargetResources | ? {$_.Type -eq "Role"}).DisplayName
            Target          = If ($event.TargetResources.Type -contains "User") {($event.TargetResources | ? {$_.Type -eq "User"}).UserPrincipalName} Else {($event.TargetResources | ? {$_.Type -eq "ServicePrincipal"}).DisplayName}
        }
    }

    $PIMevents = $PIMevents | ? {$RoleID.Keys -contains $_.Role}
}

If ($PIMevents) {
    Write-Log -Type "WRN" -Message "$(($PIMevents | Measure).Count) PIM role assignment change(s) detected in audit logs"
    
    # Send email with all the changes found in the logs
    $body = GenerateRoleChangesHTML $PIMevents
    Send-EmailAlert -Body $body -Subject "Monitor-PIMRoles - WARNING - Changes Detected"
} Else {
    Write-Log -Type "INF" -Message "No change to PIM Roles detected"
}


# Update 'Monitor-PIMRoles-SearchFrom' variable to current date/time so next time the script runs, it only searches from this point forward
Try {
    Set-AzAutomationVariable -Name "Monitor-PIMRoles-SearchFrom" -AutomationAccountName "AZU-(...)" -ResourceGroupName "AZU-(...)" -Value $newSearchFrom -Encrypted $False -ErrorAction Stop
    Write-Log -Type "INF" -Message "Variable 'Monitor-PIMRoles-SearchFrom' updated to '$newSearchFrom'"
} Catch {
    Write-Log -Type "ERR" -Message "Unable to update 'Monitor-PIMRoles-SearchFrom' variable to '$newSearchFrom': '$($_.Exception.Message)'"
    Send-EmailAlert -Body "Unable to update 'Monitor-PIMRoles-SearchFrom' variable to '$newSearchFrom': '$($_.Exception.Message)'" -Subject "Monitor-PIMRoles - ERROR - Unable to update 'Monitor-CloudShellGroup-SearchFrom' variable to '$newSearchFrom':"
}

EndScript
