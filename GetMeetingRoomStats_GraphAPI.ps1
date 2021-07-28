# Script:   GetMeetingRoomStats_GraphAPI.ps1
# Purpose:  Gather statistics regarding meeting room usage from Exchange Online
# Author:   Nuno Mota
# Date:     December 2020
# Version:  0.1 - 20200224 - First draft
#           0.2 - 20201231 - Updated to use "list places" instead of findRoom (https://docs.microsoft.com/en-us/graph/api/place-list?view=graph-rest-beta&tabs=http)
 
<#
.SYNOPSIS
Gather statistics regarding meeting room usage from Exchange Online
 
.DESCRIPTION
This script uses Graph API to connect to one or more meeting rooms and gather statistics regarding their usage between specific dates.
Although the script is targeted at meeting rooms, it will work with any mailbox default calendar in Exchange Online.
 
IMPORTANT:
    - To analyze a particular meeting room, specify one or more primary SMTP addresses in the format: 'room1@domain.com, room2@domain.com'. Alternatively, analyze all meeting rooms by using the "-All" switch;
    - You will need to have, or create, an 'app registration' in Azure and create a 'client secret';
    - The app registration will need the following API permissions to Graph API: 'User.Read.All', 'Calendars.Read', and 'Place.Read.All', all of type 'Application';
    - Maximum range to search is 1825 days (5 years);
    - You can enter the dates in the format "22/02/2020", "22/02/2020 15:00", or in ISO 8601 format such as "2020-02-22T15:00:00", or even "2020-02-22T15:00:00-08:00" to specify an offset to UTC (time zone).
 
The script gathers and exports the following stats for each meeting room for the given date range:
    - RoomName: the display name of the meeting room (when using -All). When using -RoomListSMTP, this will be the room's SMTP address;
    - RoomSMTP: the SMTP address of the meeting room;
    - From: the start of the date range to search the calendar;
    - To: the end of the date range to search the calendar;
    - totalMeetings: the total number of meetings;
    - totalDuration: the total number of minutes for all meetings;
    - totalAttendees: the total number of attendees invited across all meetings;
    - totalUniqueOrganizers: the number of unique meeting organizers;
    - totalUniqueAttendees: the number of unique attendees;
    - totalReqAttendees: the total number of required attendees;
    - totalOptAttendees: the total number of optional attendees;
    - Top5Organizers: the email address of the top 5 meeting organizers, and how many meetings each scheduled;
    - Top5Attendees: the email address of the top 5 meeting attendees, and how many meetings each attended;
    - totalAllDay: the total number of 'all-day' meetings;
    - totalAM: the total number of meetings that started in the morning (this excludes all-day meetings);
    - totalPM: the total number of meetings that started in the afternoon;
    - totalRecurring: the total number of recurring meetings;
    - totalSingle: the total number of non-recurring meetings (single instance/occurrence).
 
.PARAMETER From
    The start of the date range to search the calendar.
    You can enter the dates in the format "22/02/2020", "22/02/2020 15:00", or in ISO 8601 format such as "2020-02-22T15:00:00", or even "2020-02-22T15:00:00-08:00" to specify an offset to UTC (time zone).
 
.PARAMETER To
    The end of the date range to search the calendar.
    You can enter the dates in the format "22/02/2020", "22/02/2020 15:00", or in ISO 8601 format such as "2020-02-22T15:00:00", or even "2020-02-22T15:00:00-08:00" to specify an offset to UTC (time zone).
 
.PARAMETER All
    When using this switch, the scripts retrieves all the rooms in the tenant using the "list places" method (as oposed to using "findRooms" as in a previous version of the script;
 
.PARAMETER RoomListSMTP
    The SMTP address of one or more meeting rooms to process.
    When using this parameter, only the default calendar for the meeting room will be analyzed.
 
.PARAMETER ClientID
    The Application (client) ID of the app registration in Azure AD with required permissions.
 
.PARAMETER ClientSecret
    The secret string that the application uses to prove its identity when requesting a token. Also can be referred to as application password.
 
.PARAMETER TenantID
    The Azure Directory (tenant) ID.
 
.OUTPUTS
    The script prints to the screen the stats for each meeting room, and exports them to a CSV file in the same location of the script.
 
.LINK
    Online version: https://github.com/NunoFilipeMota/PublicScripts/blob/main/GetMeetingRoomStats_GraphAPI.ps1
 
.EXAMPLE
C:\PS> .\Get-MeetingRoomStats_GraphAPI.ps1 -All -From "01/01/2020" -To "01/02/2020"
 
Description
-----------
This command will:
   1. Process up to 100 meetings rooms in the environment;
   2. Gather statistics for all rooms for the month of January (please be aware of your date format: day/month vs month/day), using UTC format for the time;
   3. Print all stats to the screen and export them to a CSV file.
 
.EXAMPLE
C:\PS> .\GetMeetingRoomStats_GraphAPI.ps1 -RoomListSMTP "room.1@domain.com" -From "2020-02-01T00:00:00-08:00" -To "2020-03-01T00:00:00-08:00"
 
Description
-----------
This command will:
   1. Process room.1@domain.com meeting room;
   2. Gather statistics for the month of February, with a time offset of -8h compared to UTC;
   3. Print all stats to the screen and export them to a CSV file.

.EXAMPLE
C:\PS> .\GetMeetingRoomStats_GraphAPI.ps1 -RoomListSMTP "room.1@domain.com, "room.2@domain.com" -From "2020-02-01T00:00:00-08:00" -To "2020-03-01T00:00:00-08:00"
 
Description
-----------
This command will:
   1. Process the meeting rooms 'room.1@domain.com' and 'room.2@domain.com';
   2. Gather statistics for each room for the month of February, with a time offset of -8h compared to UTC;
   3. Print all stats to the screen and export them to a CSV file.

.EXAMPLE
C:\PS> Get-Help .\GetMeetingRoomStats_GraphAPI.ps1 -Full
 
Description
-----------
Shows this help manual.
#>



[CmdletBinding()]
Param (
    [Parameter(Mandatory = $False)]
    [String] $From = "2021-03-01T00:00:00",
    
    [Parameter(Mandatory = $False)]
    [String] $To = "2021-03-31T00:00:00",
    
    [Parameter(Mandatory = $False)]
    [Switch] $All,
 
    [Parameter(Mandatory = $False)]
    [String] $RoomListSMTP,
 
    [Parameter(Mandatory = $False)]
    [String] $ClientID = "",
    
    [Parameter(Mandatory = $False)]
    [String] $ClientSecret = "",
 
    [Parameter(Mandatory = $False)]
    [String] $TenantID = ""
)


#####################################################################################################
# Function to write all the actions performed by the script to a log file
#####################################################################################################
Function Write-Log {
    [CmdletBinding()]
    Param ([String] $Type, [String] $Message)
 
    $Logfile = $PSScriptRoot + "\GetMeetingRoomStats_Log_$(Get-Date -f 'yyyyMM').txt"
    If (!(Test-Path $Logfile)) {New-Item $Logfile -Force -ItemType File | Out-Null}
 
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    "$timeStamp $Type $Message" | Out-File -FilePath $Logfile -Append
 
    Switch ($Type) {
        "INF" {Write-Host $Message -ForegroundColor Green -BackgroundColor Black}
        "WRN" {Write-Host $Message -ForegroundColor Yellow -BackgroundColor Black}
        "ERR" {Write-Host $Message -ForegroundColor Red -BackgroundColor Black}
        default {Write-Host $Message}
    }
}

#####################################################################################################
# Function to get OAuth Token
#####################################################################################################
Function Get-OAuthToken {
    # Construct URI for OAuth Token and Body for OAuth Token
    $uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientID
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth Token
    Try {
        $tokenRequest = Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
        Write-Log -Type "INF" -Message "Retrieved OAuth Token"
        
        # Get token expiration date and time so we can renew it 2 minutes before it expires
        $global:tokenExpireDateTime = ((Get-Date).AddSeconds($tokenRequest.expires_in)).AddSeconds(-120)

        Return $tokenRequest.access_token
    } Catch [System.Net.WebException] {
        Write-Log -Type "ERR" -Message "Unable to get token: '$($_.Exception.Message)'"
        Exit
    }
}

#####################################################################################################
# Function to query Graph API
#####################################################################################################
Function Query-GraphAPI {
    Param ($uri, $token)
 
    # Check if we need to renew our token
    If ((Get-Date) -gt $global:tokenExpireDateTime) {$token = Get-OAuthToken}
 
    [Bool] $stopLoop = $False
    [Int32] $retryCount = 1
    Do {
        Try {
            $response = Invoke-RestMethod -Method Get -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
            $stopLoop = $True
        } Catch {
            # If we get throttled, then we sleep for 15, 30, or 45 seconds before giving up
            If ($_.Exception.Response.StatusCode -eq 429) {
                If ($retryCount -ge 3){
                    Write-Log -Type "ERR" -Message "Unable to query Graph API: '$($_.Exception.Message)'"
                    $stopLoop = $True
                    Return $False
                } Else {
                    Write-Log -Type "WRN" -Message "Unable to query Graph API: '$($_.Exception.Message)'. Retrying in $($retryCount * 15) seconds."
                    Start-Sleep -Seconds $($retryCount * 15)
                    $retryCount++
                }
            } ElseIf ($_ -match "REST API is not yet supported for this mailbox.") {
                # This error means that the meeting room hasn't been migrated to Exchange Online yet
                Write-Log -Type "WRN" -Message "$($room.nickname) has not yet been migrated to Exchange Online."
            } Else {
                Write-Log -Type "ERR" -Message "Unable to query Graph API: '$($_.Exception.Message)'"
            }

            # Write-Host $_ # Uncomment if you want to print further error details
            Return $False
        }
    } While (!$stopLoop)
 
    # Check if there are more results to retrieve (paging). The 'findRooms' method is limited to 100...
    $fullResponse = @()
    If ($response.Value) {
        $fullResponse += $response.Value
 
        If ($response."@odata.nextLink") {

            Do {
                [Bool] $stopLoop = $False
                [Int32] $retryCount = 1
                Do {
                    Try {
                        $response = Invoke-RestMethod -Uri $response."@odata.nextLink" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
                        $stopLoop = $True
                    } Catch {
                        # If we get throttled, then we sleep for 15, 30, or 45 seconds before giving up
                        If ($_.Exception.Response.StatusCode -eq 429) {
                            If ($retryCount -ge 3){
                                Write-Log -Type "ERR" -Message "Unable to query Graph API: '$($_.Exception.Message)'"
                                $stopLoop = $True
                                $fullResponse += $response.Value
                                Return $fullResponse # Return incomplete results or no results at all?
                            } Else {
                                Write-Log -Type "WRN" -Message "Unable to query Graph API: '$($_.Exception.Message)'. Retrying in $($retryCount * 15) seconds."
                                Start-Sleep -Seconds $($retryCount * 15)
                                $retryCount++
                            }
                        } ElseIf ($_ -match "REST API is not yet supported for this mailbox.") {
                            # This error means that the meeting room hasn't been migrated to Exchange Online yet
                            Write-Log -Type "WRN" -Message "$($room.name) has not yet been migrated to Exchange Online."
                        } Else {
                            Write-Log -Type "ERR" -Message "Unable to query Graph API: '$($_.Exception.Message)'"
                        }

                        # Write-Host $_ # Uncomment if you want to print further error details
                        Return $False
                    }
                } While (!$stopLoop)
 
                $fullResponse += $response.Value
            } While ($response."@odata.nextLink")
        }
    } Else {$fullResponse = $response.Value}
 
    Return $fullResponse
}

#####################################################################################################
# Script Start
#####################################################################################################
Write-Log -Type "INF" -Message "--------------------------------------------------------------------------"
Write-Log -Type "INF" -Message "START. Running under '$($env:UserName)' from '$($env:ComputerName)'."

#####################################################################################################
# Basic parameter validation
If (!$ClientID -OR !$ClientSecret -OR !$TenantID) {
    Write-Log -Type "ERR" -Message "You must use the -ClientID -ClientSecret AND -TenantID parameters. Exiting Script."
    Exit
}

If (!$All -and !$RoomListSMTP) {
    Write-Log -Type "ERR" -Message "Please use -All or -RoomListSMTP parameters. Exiting Script."
    Exit
} ElseIf ($All -and $RoomListSMTP) {
    Write-Log -Type "ERR" -Message "Please use only -All or -RoomListSMTP parameters, not both. Exiting Script."
    Exit
}

# Validate date range and convert date to ISO 8601 format, if not already in that format
If (((Get-Date $To) - (Get-Date $From)).TotalDays -gt 1825) {Write-Log -Type "ERR" -Message "The range between the start and end dates cannot be greater then 1,825 days (5 years)! Exiting script."; Exit}
Try {
    $From = Get-Date $From -Format s -ErrorAction Stop
    $To = Get-Date $To -Format s -ErrorAction Stop
} Catch {
    Write-Log -Type "ERR" -Message "Unable to convert date to ISO 8601 format: '$($_.Exception.Message)'"
    Exit
}


#####################################################################################################
# Retrieve OAuth Token
$token = Get-OAuthToken

#####################################################################################################
# If user only wants to analyse specific rooms, save the rooms details into another variable. This is
# just so we don't have to check which option was used and keep the ForEach simple
If ($RoomListSMTP) {
    [Array] $allRooms = @()
    ForEach ($room in $RoomListSMTP.Split(",") -replace (" ", "")) {
        $allRooms += [PSCustomObject] @{emailAddress = $room; nickname = $room}
    }
}

# If user selected -All, then gather all metting rooms in the tenant
If ($All) {
    # Retrieve all meeting rooms from the tenant
    # $allRooms = Query-GraphAPI -URI "https://graph.microsoft.com/beta/users/$user/findRooms" -Token $token
    $allRooms = Query-GraphAPI -URI "https://graph.microsoft.com/beta/places/microsoft.graph.room" -Token $token
    Write-Log -Type "INF" -Message "Retrieved $(($allRooms | Measure).Count) Meeting Rooms"
    If (!$allRooms) {Exit}
}


#####################################################################################################
# Gather the meetings for the selected room(s)
[Int] $count = 0
ForEach ($room in $allRooms) {
    Write-Progress -Activity "Processing Meeting Room Calendars" -Status "Processed ($("{0:N0}" -f $count) / $("{0:N0}" -f $($allRooms | Measure).Count)). Current calendar: '$($room.name)'"

    # Check if we need to renew our token
    If ((Get-Date) -gt $global:tokenExpireDateTime) {$token = Get-OAuthToken}

    # Get all room meetings for the given time period
    $allMeetings = Query-GraphAPI -URI "https://graph.microsoft.com/beta/users/$($room.emailAddress)/calendar/calendarView?startDateTime=$From&endDateTime=$To" -Token $token
    If (!$allMeetings) {Continue}
    
    $totalMeetings = ($allMeetings | Measure).Count
    If ($totalMeetings -eq 0) {
        Write-Log -Type "WRN" -Message "0 meetings retrieved from '$($room.nickname)'"
        Continue
    } Else {
        Write-Log -Type "INF" -Message "$totalMeetings meetings retrieved from '$($room.nickname)'"
    }

    [Int] $totalDuration = $totalAttendees = $totalReqAttendees = $totalOptAttendees = $totalAllDay = $totalAM = $totalPM = $totalRecurring = $totalSingle = 0
    $topOrganizers = @{}
    $topAttendees = @{}
    ForEach ($meeting in $allMeetings) {
        # Top Organizers
        $organizer = $meeting.organizer.emailAddress.address
        If ($organizer -and $topOrganizers.ContainsKey($organizer)) {
            $topOrganizers.Set_Item($organizer, $topOrganizers.Get_Item($organizer) + 1)
        } Else {
            $topOrganizers.Add($organizer, 1)
        }
 
        # Top Required Attendees
        ForEach ($attendee in ($meeting.attendees | ? {$_.Type -ne "resource"})) {
            $attendee = $attendee.emailAddress | % {$_.address}
            If (!$attendee) {Continue}
            If ($topAttendees.ContainsKey($attendee)) {
                $topAttendees.Set_Item($attendee, $topAttendees.Get_Item($attendee) + 1)
            } Else {
                $topAttendees.Add($attendee, 1)
            }
        }

        # Gather other stats
        $totalDuration += ((Get-Date $meeting.end.dateTime) - (Get-Date $meeting.start.dateTime)).TotalMinutes
        $totalAttendees += ($meeting.attendees | ? {$_.Type -ne "resource"} | Measure).Count
        $totalReqAttendees += ($meeting.attendees | ? {$_.Type -eq "required"} | Measure).Count
        $totalOptAttendees += ($meeting.attendees | ? {$_.Type -eq "optional"} | Measure).Count
        If ($meeting.isAllDay) {$totalAllDay++} Else {If ((Get-Date $meeting.start.dateTime -UFormat %p) -eq "AM") {$totalAM++} Else {$totalPM++}}
        If ($meeting.type -eq "occurrence") {$totalRecurring++} Else {$totalSingle++}

        # # If you want to capture the details of each meeting
        # [PSCustomObject] @{
        #   Start       = (Get-Date $meeting.start.dateTime).ToString() + " " + $meeting.start.timeZone
        #   End         = (Get-Date $meeting.end.dateTime).ToString() + " " + $meeting.end.timeZone
        #   Subject     = $meeting.subject
        #   importance  = $meeting.importance
        #   isAllDay    = $meeting.isAllDay
        #   MeetingRoom = $meeting.locations.locationUri -Join "; "
        #   attendees   = ($meeting.attendees.emailAddress | % {$_.address}) -Join "; "
        #   organizer   = $meeting.organizer.emailAddress.address
        #   type        = $meeting.type
        #   hasAttachments  = $meeting.hasAttachments
        # }
    }

    $roomsObj = [PSCustomObject] @{
        RoomName                = $room.nickname
        RoomSMTP                = $room.emailAddress
        From                    = Get-Date $From
        To                      = Get-Date $To
        totalMeetings           = "{0:N0}" -f $totalMeetings
        totalDuration           = "{0:N0}" -f $totalDuration
        totalAttendees          = "{0:N0}" -f $totalAttendees
        totalUniqueOrganizers   = If ($topOrganizers) {($topOrganizers.GetEnumerator() | Select Name | Measure).Count} Else {""}
        totalUniqueAttendees    = If ($topAttendees) {($topAttendees.GetEnumerator() | Select Name | Measure).Count} Else {""}
        totalReqAttendees       = $totalReqAttendees
        totalOptAttendees       = $totalOptAttendees
        Top5Organizers          = If ($topOrganizers) {($topOrganizers.GetEnumerator() | Sort -Property Value -Descending | Select -First 5 | % {"$($_.Key) ($($_.Value))"}) -Join ", "} Else {""}
        Top5Attendees           = If ($topAttendees) {($topAttendees.GetEnumerator() | Sort -Property Value -Descending | Select -First 5 | % {"$($_.Key) ($($_.Value))"}) -Join ", "} Else {""}
        totalAllDay             = "{0:N0}" -f $totalAllDay
        totalAM                 = "{0:N0}" -f $totalAM
        totalPM                 = "{0:N0}" -f $totalPM
        totalRecurring          = "{0:N0}" -f $totalRecurring
        totalSingle             = "{0:N0}" -f $totalSingle
    }

    $roomsObj
    $roomsObj | Export-CSV "GetMeetingRoomStats_$(Get-Date -f 'yyyyMMdd').csv" -NoType -Append
    $count++
}

Write-Log -Type "INF" -Message "END."
