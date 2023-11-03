# Script:   Distribution Group Stats.ps1
# Purpose:  Collects statistics about Distribution Groups and uploads them to a SharePoint Online List
# Author:   Nuno Mota
# Date:     July 2022
# Version:  0.1 - 202207 - First draft
#           0.2 - 202306 - Updated to Get-EXORecipient because of https://learn.microsoft.com/en-us/answers/questions/1023619/connect-exchangeonline-write-errormessage-expired

<#
.SYNOPSIS
Collects statistics about Distribution Groups and uploads them to a SharePoint Online List
 
.DESCRIPTION
This script uses Exchange Online PowerShell module to collect statistics about Distribution Groups and then uses Graph API to upload them to a SharePoint Online List
 
IMPORTANT:
    - This script is not really suitable for Azure Automation given the 3h limit per runbook (unless using hybrid runbook worker, or limiting the number of DGs processed each time - check code line ?????????????????????????????);
	- The script is slow... In my environment it takes around 2-3 seconds to process each DG;
    - Message Tracking is only available for the past 10 days (unless a CSV output is used). As such, the script will have to run for a whole month until you see accurate results in terms of emails sent/received;
    - You should run the script frequently enough so it updates all DGs at least once every 10 days (so you don't miss any sent/received emails);
    - Please use PowerShell 7+. PowerShell 5.1 default memory allocation can cause the script to crash when fetching large data sets;
	- The script uses Graph API SDK, so you will need the 'Microsoft.Graph.Sites' and 'Microsoft.Graph.Users' modules;
		- More information on the SDK: https://docs.microsoft.com/en-us/graph/powershell/get-started
	- You will need to have, or create, an 'app registration' in Azure and use a digital certificate for authentication (for both Exchange and SharePoint access):
		- Use app-only authentication with the Microsoft Graph PowerShell SDK: https://docs.microsoft.com/en-us/graph/powershell/app-only?tabs=azure-portal
	- The Graph API permissions required for the script to work are 'Sites.ReadWrite.All' ('Sites.Selected' instead will also work) and 'Mail.Send', plus Exchange Online;

.OUTPUTS
    1. The script prints to the screen any errors as well as all successful actions taken;
	2. It also generates a log file named "DistributionGroupStats_Log_<Date>" with the same information;
	3. Any new DGs found are added to the specified SharePoint List;
	4. Any existing DGs (after the script has run at least once) are updated in the specified SharePoint List.
	5. Any deleted DGs (after the script has run at least once) are deleted from the specified SharePoint List.
 
.LINK
    Online version: https://github.com/NunoFilipeMota/PublicScripts/blob/main/DistributionGroupStats.ps1
 
.EXAMPLE
C:\PS> .\DistributionGroupStats.ps1
 
Description
-----------
This command will:
   1. Retrieve all the DGs that have been added, if any, to the SharePoint List;
   2. If there are any DGs retrieved from the SharePoint list, update their details, including sent/received emails. If any no longer exist, they are deleted from the list;
   3. Retrieve all DGs from Exchange Online in order to add any newly created ones to the SharePoint list.
 
.EXAMPLE
C:\PS> Get-Help .\DistributionGroupStats.ps1 -Full
 
Description
-----------
Shows this help manual.
#>



#####################################################################################################
# Function to write all the actions performed by the script to a log file
#####################################################################################################
Function Write-Log {
    [CmdletBinding()]
    Param ([String] $Type, [String] $Message)

    $Logfile = "$PSScriptRoot\DistributionGroupStats_Log_$(Get-Date -f 'yyyyMM').txt"
    If (!(Test-Path $Logfile)) {New-Item $Logfile -Force -ItemType File | Out-Null}

    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    "$timeStamp $Type $Message" | Out-File -FilePath $Logfile -Append

    Switch ($Type) {
       "INF" {Write-Host $Message -ForegroundColor Green -BackgroundColor Black}
       "WRN" {Write-Host $Message -ForegroundColor Yellow -BackgroundColor Black}
       "ERR" {Write-Host $Message -ForegroundColor Red -BackgroundColor Black}
       default {Write-Host $Message}
    }

    # Use this instead if using Azure Automation
    # Switch ($Type) {
    #     "INF" {Write-Output $Message}
    #     "WRN" {Write-Warning $Message}
    #     "ERR" {Write-Error $Message}
    #     default {$Message}
    # }
}


#####################################################################################################
# Function to connect to Exchange Online
#####################################################################################################
Function ConnectExchangeOnline {
    If (!(Get-PSSession | ? {$_.ComputerName -eq "outlook.office365.com" -AND $_.State -eq "Opened"})) {
        Try {
            Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $ClientID -Organization "domain.onmicrosoft.com" -CommandName Get-DistributionGroup, Get-DistributionGroupMember, Get-RecipientPermission, Get-EXORecipient, Get-Recipient, Get-MessageTrace -ErrorAction Stop
            Write-Log -Type "INF" -Message "Connected to Exchange Online."
        } Catch {
            Write-Log -Type "ERR" -Message "Unable to establish PowerShell session to Exchange Online: '$($_.Exception.Message)'. Exiting script"
            Send-EmailAlert -Body "Unable to establish PowerShell session to Exchange Online: '$($_.Exception.Message)'. Exiting script" -Subject "Distribution Group Stats - ERROR - Unable to establish PowerShell session to Exchange Online"
            EndScript
        }
    }
}


#####################################################################################################
# Function to connect to Graph API
#####################################################################################################
Function ConnectGraphAPI {
    Try {
        Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
        Import-Module Microsoft.Graph.Sites -ErrorAction Stop
        Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -CertificateThumbprint $CertThumbprint -ErrorAction Stop
        Write-Log -Type "INF" -Message "Connected to Graph API"
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to connect to Graph API: '$($_.Exception.Message)'"
        # Send-EmailAlert -Body "Unable to connect to Graph API: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to connect to Graph API"
        EndScript
    }
}


#####################################################################################################
# Function to send an email alert when something goes wrong
#####################################################################################################
Function Send-EmailAlert {
    Param ([String] $Body, [String] $Subject)

    $postbody = @{
        Message = @{
            Importance = "High"
            Subject = $Subject
            Body = @{
                ContentType = "HTML"
                Content = $Body
            }
            ToRecipients = @(
                @{EmailAddress = @{Address = "nuno.mota@domain.com"}}
            )
        }
        saveToSentItems = $False
    }

    Try {
        Send-MgUserMail -UserId "nuno.mota@domain.com" -BodyParameter $postbody -ErrorAction Stop
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to send alert email: '$($_.Exception.Message)'"
    }
}


#####################################################################################################
# Function to upload (or update) DG details to SharePoint list
#####################################################################################################
Function UploadUpdate-SharePointList {
    Param ($DG)

    $listParams = @{
        DName                       = $DG.DName; `
        PrimarySmtpAddress          = $DG.PrimarySmtpAddress; `
        Description                 = $DG.Description; `
        SentEmailsThisMonth         = $DG.SentEmailsThisMonth; `
        SentEmailsLastMonth         = $DG.SentEmailsLastMonth; `
        SentEmailsMonthBefore       = $DG.SentEmailsMonthBefore; `
        ReceivedEmailsThisMonth     = $DG.ReceivedEmailsThisMonth; `
        ReceivedEmailsLastMonth     = $DG.ReceivedEmailsLastMonth; `
        ReceivedEmailsMonthBefore   = $DG.ReceivedEmailsMonthBefore; `
        IsDirSynced                 = $DG.IsDirSynced; `
        OwnersCount                 = $DG.OwnersCount; `
        Owners                      = $DG.Owners; `
        MemberCount                 = $DG.MemberCount; `
        Members                     = $DG.Members; `
        JoinRestriction             = $DG.JoinRestriction; `
        DepartRestriction           = $DG.DepartRestriction; `
        HiddenFromGAL               = $DG.HiddenFromGAL; `
        RecipientTypeDetails        = $DG.RecipientTypeDetails
        SecurityGroup               = $DG.SecurityGroup; `
        SenderAuth                  = $DG.SenderAuth; `
        SendAs                      = $DG.SendAs; `
        SendOnBehalf                = $DG.SendOnBehalf; `
        WhenChangedUTC              = $DG.WhenChangedUTC; `
        WhenCreatedUTC              = $DG.WhenCreatedUTC; `
        LastUpdated                 = $DG.LastUpdated; `
        DaysOld                     = $DG.DaysOld
    }

    # If we have a ListItemID, then we are updating a record, if not, we are creating a new one
    If ($DG.ListItemID) {
        Try {
            Update-MgSiteListItem -SiteId $SiteID -ListId $ListID -ListItemId $DG.ListItemID -Fields $listParams | Out-Null #-ErrorAction Stop
            Write-Log -Type "INF" -Message "Updated '$($DG.PrimarySmtpAddress)'"
        } Catch {
            Write-Log -Type "ERR" -Message "Unable to update '$($DG.PrimarySmtpAddress)' in SharePoint List: '$($_.Exception.Message)'"
            Send-EmailAlert -Body "Unable to update '$($DG.PrimarySmtpAddress)' in SharePoint List: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to update '$($DG.PrimarySmtpAddress)' in SharePoint List"
        }
    } Else {
        Try {
            New-MgSiteListItem -SiteId $siteID -ListId $listID -Fields $listParams -ErrorAction Stop | Out-Null
            Write-Log -Type "INF" -Message "Created '$($DG.PrimarySmtpAddress)'"
        } Catch {
            Write-Log -Type "ERR" -Message "Unable to create '$($DG.PrimarySmtpAddress)' entry in SharePoint List: '$($_.Exception.Message)'"
            Write-Log -Type "ERR" -Message $listParams
            Write-Log -Type "ERR" -Message $($listParams | Out-String)
            Send-EmailAlert -Body "Unable to create '$($DG.PrimarySmtpAddress)' in SharePoint List: '$($_.Exception.Message)'.`n`ns $($listParams | Out-String)" -Subject "Distribution Group Stats - ERROR - Unable to create '$($DG.PrimarySmtpAddress)' in SharePoint List"
        }
    }
}


#####################################################################################################
# Function to get information about a Distribution Group
#####################################################################################################
Function Get-DGinfo {
    Param ($smtp, $existingDG)

    # Check if we are processign a DG for the first time or an existing DG
    If ($smtp) {$email = $smtp} Else {$email = $existingDG.PrimarySmtpAddress}

    # Write-Log -Type "INF" -Message "Retrieving details for DG '$email'"
    Try {
        # Using Filter in case there's an Microsoft 365 with the same address/details. They both would have MailUniversalDistributionGroup as RecipientType, but MailUniversalDistributionGroup/GroupMailbox as RecipientTypeDetails
        $DG = Get-DistributionGroup $email -Filter {RecipientTypeDetails -ne "GroupMailbox"} -ErrorAction Stop
    } Catch {
        Switch -Wildcard ($_.Exception.Message) {
            # Check the error for non-existing mailboxes (i.e., deleted)
            "*couldn't be found on*" {
                Try {
                    Remove-MgSiteListItem -SiteId $SiteID -ListId $ListID -ListItemId $existingDG.ListItemID -ErrorAction Stop | Out-Null
                    Write-Log -Type "INF" -Message "Deleted '$($existingDG.PrimarySmtpAddress)' from SharePoint List"
                    $global:countDeleted++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to delete '$($existingDG.PrimarySmtpAddress)' from SharePoint List: '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to delete '$($existingDG.PrimarySmtpAddress)' from SharePoint List: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to delete '$($existingDG.PrimarySmtpAddress)' from SharePoint List"
                }
            }

            # Check if we are being throttled. If yes, sleep for the suggested ammount of time
            "*This operation exceeds the throttling budget for policy*" {
                $backoff = $_.Exception.Message | Select-String -Pattern "[\d]* ms" -AllMatches
                If ($backoff.Matches.Value) {
                    $backoff = [Int]$backoff.Matches.Value.Split(" ms")[0] + 2000
                    Write-Log -Type "WRN" -Message "Unable to retrieve '$email' DG because of throttling. Sleeping for $backoff ms"
                    Send-EmailAlert -Body "Unable to retrieve '$email' DG because of throttling. Sleeping for $backoff ms" -Subject "Distribution Group Stats - ERROR - Unable to retrieve '$email' DG because of throttling. Sleeping for $backoff ms"
                    Start-Sleep -Milliseconds $backoff
                }
            }

            "*is not recognized as a name of a cmdlet, function, script file, or executable program*" {
                Write-Log -Type "ERR" -Message "Failed to reconnect to Exchange Online. Exiting script."
                Send-EmailAlert -Body "Failed to reconnect to Exchange Online. Exiting script." -Subject "Distribution Group Stats - ERROR - Failed to reconnect to Exchange Online - Exiting script"
                EndScript
            }

            "*Starting a command on the remote server failed with the following error message*" {
                Write-Log -Type "ERR" -Message "Failed to reconnect to Exchange Online. Exiting script."
                Send-EmailAlert -Body "Failed to reconnect to Exchange Online. Exiting script." -Subject "Distribution Group Stats - ERROR - Failed to reconnect to Exchange Online - Exiting script"
                EndScript
            }

            Default {
                Write-Log -Type "ERR" -Message "Unable to retrieve '$email' DG: '$($_.Exception.Message)'. Skipping DG."
                
                # Not sending email alerts for this specific issue as it might happen for a couple hundred DGs when analysing thousands of DGs
                If ($_.Exception.Message -notmatch "An error caused a change in the current set of domain controllers" -AND $_.Exception.Message -notmatch "An error caused a change in the current set of domain controllers") {
                    Send-EmailAlert -Body "Unable to retrieve '$email' DG: '$($_.Exception.Message)'. Skipping DG." -Subject "Distribution Group Stats - ERROR - Unable to retrieve '$email' DG"
                }
            }
        }

        Return $False
    }


    #####################################################################################################
    # Retrieve Owners (in case there are many that will take the resulting string over the 255 char limit in SharePoint)
    $OwnerEmails = ($DG.ManagedBy | Get-Recipient -ErrorAction SilentlyContinue).PrimarySMTPAddress -Join ";"
    If ($OwnerEmails.Length -gt 255) {$OwnerEmails = $OwnerEmails.Substring(0, 249) + "(...)"}

    #####################################################################################################
    # Retrieve Members
    Try {
        $Members = ""
        $Members = (Get-DistributionGroupMember $DG.DistinguishedName -ResultSize Unlimited -ErrorAction Stop | ? {$_.PrimarySmtpAddress.ToString()}).PrimarySmtpAddress -Join ";"
        If ($Members.Length -gt 255) {$Members = $Members.Substring(0, 249) + "(...)"}
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to retrieve DG members: '$($_.Exception.Message)'"
        Send-EmailAlert -Body "Unable to retrieve '$($DG.PrimarySMTPAddress)' DG members: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve '$($DG.PrimarySMTPAddress)' DG members"
    }

    #####################################################################################################
    # Retrieve SendAs permissions
    Try {
        $SendAsEmails = ""
        $SendAs = (Get-RecipientPermission $DG.PrimarySmtpAddress -AccessRights SendAs -WarningAction:SilentlyContinue -ErrorAction Stop | ? {!$_.IsInherited}).Trustee
        If ($SendAs) {
            $SendAsEmails = $SendAs -Join ";"
            If ($SendAsEmails.Length -gt 255) {$SendAsEmails = $SendAsEmails.Substring(0, 249) + "(...)"}
        }
    } Catch {
        Write-Log -Type "ERR" -Message "Unable to retrieve SendAs permissions: '$($_.Exception.Message)'"
        Send-EmailAlert -Body "Unable to retrieve '$($DG.PrimarySMTPAddress)' SendAs permissions: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve '$($DG.PrimarySMTPAddress)' SendAs permissions"
    }

    # Retrieve Send-On-Behalf permissions
    $SendOnBehalfEmails = ""
    $SendOnBehalf = ($DG.GrantSendOnBehalfTo | Get-Recipient -ErrorAction SilentlyContinue).PrimarySMTPAddress
    If ($SendOnBehalf) {
        $SendOnBehalfEmails = $SendOnBehalf -Join ";"
        If ($SendOnBehalfEmails.Length -gt 255) {$SendOnBehalfEmails = $SendOnBehalfEmails.Substring(0, 249) + "(...)"}
    }

    # Construct a PowerShell object with the DG details and stats
    $dgObj = [PSCustomObject] [Ordered] @{
        ListItemID                  = If ($smtp) {$null} Else {$existingDG.ListItemID}
        DName                       = $DG.DisplayName
        PrimarySmtpAddress          = $DG.PrimarySmtpAddress
        Description                 = If ($DG.Description.Length -gt 255) {$($DG.Description.Substring(0, 249) + "(...)")} Else {$DG.Description}
        SentEmailsThisMonth         = If ($smtp) {0} Else {$existingDG.SentEmailsThisMonth}         # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        SentEmailsLastMonth         = If ($smtp) {0} Else {$existingDG.SentEmailsLastMonth}         # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        SentEmailsMonthBefore       = If ($smtp) {0} Else {$existingDG.SentEmailsMonthBefore}       # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        ReceivedEmailsThisMonth     = If ($smtp) {0} Else {$existingDG.ReceivedEmailsThisMonth}     # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        ReceivedEmailsLastMonth     = If ($smtp) {0} Else {$existingDG.ReceivedEmailsLastMonth}     # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        ReceivedEmailsMonthBefore   = If ($smtp) {0} Else {$existingDG.ReceivedEmailsMonthBefore}   # If we are processing a DG for the first time, it won't have any sent/rec email information at this stage
        IsDirSynced                 = $DG.IsDirSynced
        OwnersCount                 = If ($OwnerEmails) {($OwnerEmails -Split ";" | Measure).Count} Else {0}
        Owners                      = $OwnerEmails
        MemberCount                 = If ($Members) {($Members -Split ";" | Measure).Count} Else {0}
        Members                     = $Members
        JoinRestriction             = $DG.MemberJoinRestriction
        DepartRestriction           = $DG.MemberDepartRestriction
        HiddenFromGAL               = $DG.HiddenFromAddressListsEnabled
        RecipientTypeDetails        = $DG.RecipientTypeDetails
        SecurityGroup               = $DG.RecipientTypeDetails -match "Security"
        SenderAuth                  = $DG.RequireSenderAuthenticationEnabled
        SendAs                      = $SendAsEmails
        SendOnBehalf                = $SendOnBehalfEmails
        WhenChangedUTC              = (Get-Date $DG.WhenChangedUTC).ToString("yyyy-MM-ddTHH:mm:ssZ")
        WhenCreatedUTC              = (Get-Date $DG.WhenCreatedUTC).ToString("yyyy-MM-ddTHH:mm:ssZ")
        LastUpdated                 = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        DaysOld                     = ((Get-Date) - (Get-Date $DG.WhenCreatedUTC)).Days
    }


    #region SentReceivedEmails
    #####################################################################################################
    # Sent and Received Emails
    # We still search sent emails for DGs without SendAs or SentOnBehalf permissions as those permissions could have been configured in the past 10 days
    $pageSize = 5000
    $endDate = Get-Date

    # Check if we are processign an existing DG, or a new DG for the first time (in which case we can only search the last 10 days)
    If ($smtp) {
        # Check the last 10 days' worth of emails. If the search window expands two different months, just get this months' emails
        $startDate = (Get-Date).AddDays(-10)
        If ((Get-Date $startDate -f "MMMM") -ne (Get-Date $endDate -f "MMMM")) {$startDate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0}

        # Received Emails
        $page = 1 
        Do {
            Try {
                $currentSearch = Get-MessageTrace -RecipientAddress $DG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Expanded -PageSize $pageSize -Page $page | Select MessageID
                $dgObj.ReceivedEmailsThisMonth += ($currentSearch | Measure).Count
                $page++
            } Catch {
                Write-Log -Type "ERR" -Message "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                Send-EmailAlert -Body "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)'"
                Break
            }
        } Until ($currentSearch.Count -lt $pageSize)

        # Sent Emails
        $page = 1 
        Do {
            Try {
                $currentSearch = Get-MessageTrace -SenderAddress $DG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Delivered -PageSize $pageSize -Page $page | Select MessageID
                $dgObj.SentEmailsThisMonth += ($currentSearch | Select -Unique MessageID | Measure).Count
                $page++
            } Catch {
                Write-Log -Type "ERR" -Message "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                Send-EmailAlert -Body "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to sent received emails for '$($DG.PrimarySmtpAddress)'"
                Break
            }
        } Until ($currentSearch.Count -lt $pageSize)
    } Else {
        # We are processing an existing DG, therefore with email data already in SharePoint
        # Check if the last time we searched for emails was more than 10 days ago or not
        If (((Get-Date) - (Get-Date $existingDG.LastUpdated)).TotalDays -gt 10) {
            $startDate = (Get-Date).AddDays(-10)
        } Else {
            $startDate = Get-Date $existingDG.LastUpdated
        }

        # Check if the search window expands two different months
        If ((Get-Date $startDate -f "MMMM") -eq (Get-Date -f "MMMM")) {
            # We are searching for emails in the same month
            # Received Emails
            $page = 1
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -RecipientAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Expanded -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.ReceivedEmailsThisMonth += ($currentSearch | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)

            # Sent Emails
            $page = 1 
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -SenderAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Delivered -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.SentEmailsThisMonth += ($currentSearch | Select -Unique MessageID | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)
        } Else {
            # We're in a new month, so we need to update last month's emails, and then get this month's emails for the first time
            $tempEnddate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0

            #####################################################################################################
            # Received Emails - last month
            $page = 1 
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -RecipientAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $tempEnddate -Status Expanded -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.ReceivedEmailsThisMonth += ($currentSearch | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)
            $dgObj.ReceivedEmailsMonthBefore = $dgObj.ReceivedEmailsLastMonth
            $dgObj.ReceivedEmailsLastMonth = $dgObj.ReceivedEmailsThisMonth
            $dgObj.ReceivedEmailsThisMonth = 0

            # Sent Emails - last month
            $page = 1 
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -SenderAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $tempEnddate -Status Delivered -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.SentEmailsThisMonth += ($currentSearch | Select -Unique MessageID | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to sent received emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)
            $dgObj.SentEmailsMonthBefore = $dgObj.SentEmailsLastMonth
            $dgObj.SentEmailsLastMonth = $dgObj.SentEmailsThisMonth
            $dgObj.SentEmailsThisMonth = 0

            #####################################################################################################
            # This month
            $startDate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0

            # Received Emails - this month
            $page = 1 
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -RecipientAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Expanded -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.ReceivedEmailsThisMonth += ($currentSearch | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve received emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)

            # Sent Emails - this month
            $page = 1 
            Do {
                Try {
                    $currentSearch = Get-MessageTrace -SenderAddress $existingDG.PrimarySmtpAddress -StartDate $startDate -EndDate $endDate -Status Delivered -PageSize $pageSize -Page $page | Select MessageID
                    $dgObj.SentEmailsThisMonth += ($currentSearch | Select -Unique MessageID | Measure).Count
                    $page++
                } Catch {
                    Write-Log -Type "ERR" -Message "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'"
                    Send-EmailAlert -Body "Unable to retrieve sent emails for '$($DG.PrimarySmtpAddress)': '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to sent received emails for '$($DG.PrimarySmtpAddress)'"
                    Break
                }
            } Until ($currentSearch.Count -lt $pageSize)
        }
    }
    #endregion SentReceivedEmails

    Return $dgObj
}


#####################################################################################################
# Function to terminate the script
#####################################################################################################
Function EndScript {
    Try {Disconnect-ExchangeOnline -Confirm:$False} Catch {}
    Try {Disconnect-MgGraph -ErrorAction Stop} Catch {}
    Write-Log -Type "INF" -Message "DGs Updated: $('{0:N0}' -f $countUpdated) | DGs added: $('{0:N0}' -f $countAdded) | DGs deleted: $('{0:N0}' -f $global:countDeleted)"
    $stopWatch.Stop()
    Write-Log -Type "INF" -Message "END (runtime: $('{0:N0}' -f $($stopWatch.Elapsed.TotalMinutes)) minutes / $('{0:N1}' -f $($stopWatch.Elapsed.TotalHours)) hours)"
    Exit
}


#####################################################################################################
# START
$stopWatch = [System.Diagnostics.Stopwatch]::startNew()
Write-Log -Type "INF" -Message "--------------------------------------------------------------------------"
Write-Log -Type "INF" -Message "START"

# Initialise variables used throughout the script
$ClientID = ""
$TenantID = ""
$CertThumbprint = ""
$SiteID = "domain.sharepoint.com,c7e65482-e35e-46d2-8d09-e7f69ccf2f54,735ceffc-1e1c-4dcf-b78d-d0095f25f65b"
$ListID = "6548f43c-6514-6546-a945-5430554ac2e6"
$count = $countTotal = $countUpdated = $countAdded = $global:countDeleted = 0

# Connect to Graph API and Exchange Online
ConnectExchangeOnline
ConnectGraphAPI


#####################################################################################################
# Retrieve DG items from SharePoint List (https://..........)
Try {
    $listItems = Get-MgSiteListItem -SiteId $SiteID -ListId $ListID -All -ExpandProperty "fields"
    
    $countTotal = ($listItems | measure).Count
    If ($countTotal -eq 0) {
        Write-Log -Type "INF" -Message "No DGs to update in SharePoint list"
    } Else {
        Write-Log -Type "INF" -Message "$("{0:N0}" -f $countTotal) DG(s) retrieved from SharePoint list."
        $dgCol = @()
    }
} Catch {
    Write-Log -Type "ERR" -Message "Unable to retrieve items from SharePoint List: '$($_.Exception.Message)'"
    Send-EmailAlert -Body "Unable to retrieve items from SharePoint List: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve items from SharePoint List"
}


# Get all the DGs from the SharePoint List so we can updat them all. Creating an array of objects with their details to make it easier to work with
ForEach ($item in $listItems) {
    # Get DG details from the SharePoint list
    $DG = $item.fields.AdditionalProperties

    $dgCol += [PSCustomObject] [Ordered] @{
        ListItemID                  = $item.Id
        DName                       = $DG.DName
        PrimarySmtpAddress          = $DG.PrimarySmtpAddress
        Description                 = $DG.Description
        SentEmailsThisMonth         = $DG.SentEmailsThisMonth
        SentEmailsLastMonth         = $DG.SentEmailsLastMonth
        SentEmailsMonthBefore       = $DG.SentEmailsMonthBefore
        ReceivedEmailsThisMonth     = $DG.ReceivedEmailsThisMonth
        ReceivedEmailsLastMonth     = $DG.ReceivedEmailsLastMonth
        ReceivedEmailsMonthBefore   = $DG.ReceivedEmailsMonthBefore
        IsDirSynced                 = $DG.IsDirSynced
        OwnersCount                 = $DG.OwnersCount
        Owners                      = $DG.Owners
        MemberCount                 = $DG.MemberCount
        Members                     = $DG.Members
        JoinRestriction             = $DG.JoinRestriction
        DepartRestriction           = $DG.DepartRestriction
        HiddenFromGAL               = $DG.HiddenFromGAL
        RecipientTypeDetails        = $DG.RecipientTypeDetails
        SecurityGroup               = $DG.SecurityGroup
        SenderAuth                  = $DG.SenderAuth
        SendAs                      = $DG.SendAs
        SendOnBehalf                = $DG.SendOnBehalf
        WhenChangedUTC              = $DG.WhenChangedUTC
        WhenCreatedUTC              = $DG.WhenCreatedUTC
        LastUpdated                 = $DG.LastUpdated
        DaysOld                     = $DG.DaysOld
    }
}

# Update DGs in the SharePoint List. Here we can update all of them, or the x DGs that haven't been updated in the longest.
# I suggest running the script often with a small number (4000 or less) to avoid 'An error caused a change in the current set of domain controllers' errors
If ($dgCol) {
    # ForEach ($DG in $dgCol) {
    ForEach ($DG in ($dgCol | Sort LastUpdated | Select -First 4000)) {
        $DGdetails = Get-DGinfo $null $DG

        If ($DGdetails) {
            UploadUpdate-SharePointList $DGdetails
            $countUpdated++
        }

        $count++
	    Write-Progress -Activity "Processing DGs in SharePoint List" -Status "Processed $("{0:N0}" -f $count) / $("{0:N0}" -f $countTotal)"
    }
}


#####################################################################################################
# Retrieve all Distribution Groups so we can create entries for newly created ones
Try {
    Write-Log -Type "INF" -Message "Retrieving all Distribution Groups from Exchange Online"
    # $DGs = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup -ErrorAction Stop | Select PrimarySmtpAddress, EmailAddresses    # https://learn.microsoft.com/en-us/answers/questions/1023619/connect-exchangeonline-write-errormessage-expired
    $DGs = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup -ErrorAction Stop | Select PrimarySmtpAddress, EmailAddresses

    $count = 0
    $countTotal = ($DGs | Measure).Count
    Write-Log -Type "INF" -Message "$("{0:N0}" -f $countTotal) DGs retrieved"
} Catch {
    Write-Log -Type "ERR" -Message "Unable to retrieve DGs from Exchange Online: '$($_.Exception.Message)'. Ending script"
    Send-EmailAlert -Body "Unable to retrieve DGs from Exchange Online: '$($_.Exception.Message)'" -Subject "Distribution Group Stats - ERROR - Unable to retrieve DGs from Exchange Online. Ending script"
    EndScript
}

# Process all the DGs retrieved from Exchange Online, and create entries for those that are not already in the SharePoint List
:labelA ForEach ($DG in $DGs) {
    ForEach ($email in $DG.EmailAddresses) {
        # Using this method instead of simply comparing $dgCol.PrimarySmtpAddress with $DG.PrimarySmtpAddress in case the primary address of a DG has been updated (in which case a new entry would have been created)
        If ($dgCol.PrimarySmtpAddress -contains $($email -Replace ("smtp:", ""))) {
            $count++
            Write-Progress -Activity "Processing all shared mailboxes in Exchange Online" -Status "Processed $("{0:N0}" -f $count) / $("{0:N0}" -f $countTotal)"
            Continue labelA
        }
    }

    $DGdetails = Get-DGinfo $DG.PrimarySmtpAddress $null
    If ($DGdetails) {
        UploadUpdate-SharePointList $DGdetails
        $countAdded++
    }

    $count++
    Write-Progress -Activity "Processing all DGs in Exchange Online" -Status "Processed $("{0:N0}" -f $count) / $("{0:N0}" -f $countTotal)"
}


EndScript
