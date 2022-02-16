# Script:  CreateCalendarEvent.ps1
# Purpose: Create a calendar event on users' Exchange Online mailboxes
# Author:  Nuno Mota
# Date:    February 2022
# Version: 1.0 - 202202 - First release
#          1.1 - 202202 - Updated notes about Hybrid environments

<#
.SYNOPSIS
Create a calendar event on users' Exchange Online mailboxes
 
.DESCRIPTION
This script uses Graph API to create a calendar event on users' mailboxes hosted in Exchange Online.
 
IMPORTANT:
    - If you have many users in Azure AD, please use PowerShell 7+. PowerShell 5.1 default memory allocation will cause the script to crash when fetching large data sets;
	- The script uses Graph API SDK, so you will need the 'Microsoft.Graph.Calendar' and 'Microsoft.Graph.Users' (if using the -AllUsers switch) modules;
		- More information on the SDK: https://docs.microsoft.com/en-us/graph/powershell/get-started
	- You will need to have, or create, an 'app registration' in Azure and use a digital certificate for authentication:
		- Use app-only authentication with the Microsoft Graph PowerShell SDK: https://docs.microsoft.com/en-us/graph/powershell/app-only?tabs=azure-portal
	- The Graph API permissions required for the script to work are 'Calendars.ReadWrite' and 'User.Read.All' (if using the -AllUsers switch). Both of type Application;
		- Create Event: https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0&tabs=http
	- Whenever the script successfully creates an event on a user's mailbox, it saves the user's SMTP/UPN to a file named 'CreateCalendarEvent_Processed.txt'. This is so the file can be used to re-run the script for any remaining users (in case of a timeout or any other issues).
	- The script is slow... When I ran it for 37000+ users, it took approximately 1 second per user. Need to look into JSON batching (https://docs.microsoft.com/en-us/graph/json-batching).
	- The script will throw errors in an Hybrid environment with mailboxes on-prem (as they are returned by Get-MgUser). If this is your case, you might want to use an Exchange Online cmdlet instead of Get-MgUser (or get all your mailboxes and then use the -UsersFile parameter).
 
.PARAMETER UsersFile
    TXT file containing the email addresses or UPNs of the mailboxes to create a calendar event on.
 
.PARAMETER ExcludeUsersFile
    TXT file containing the email addresses of the mailboxes NOT to create a calendar event on.
 
.PARAMETER AllUsers
    Creates a calendar event on all Exchange Online mailboxes of enabled users that have an EmployeeID configured.
	This can, and should, be adapted to your specific environment or requirement.
	The script does not use Exchange Online to retrieve the list of mailboxes. It retrieves all users from Azure AD that have the Mail and EmployeeID attributes populated.
 
.OUTPUTS
    1. The script prints to the screen any errors as well as all successful calendar entries created.
	2. It also generates a log file named "CreateCalendarEvent_Log_<Date>" with the same information.
	3. Whenever it successfully creates an event on a user's mailbox, it outputs the user's SMTP/UPN to a file named "CreateCalendarEvent_Processed.txt". This is so the file can be used to re-run the script for any remaining users (in case of a timeout or any other issues) without the risk of duplciating calendar entries.
	4. For any failures when creating a calendar event, the script writes the user's SMTP/UPN to a file named "CreateCalendarEvent_Failed.txt" so admins can easily analyse failures (the same is written to the main log file).
 
.LINK
    Online version: https://github.com/NunoFilipeMota/PublicScripts/blob/main/CreateCalendarEvent.ps1
 
.EXAMPLE
C:\PS> .\CreateCalendarEvent.ps1 -AllUsers
 
Description
-----------
This command will:
   1. Retrieve all users from Azure AD that have the Mail and EmployeeID attributes populated;
   2. Create a calendar event on their mailboxes. The properties of the calendar event are detailed and configurable within the script.
 
.EXAMPLE
C:\PS> .\CreateCalendarEvent.ps1 -AllUsers -ExcludeUsersFile .\CreateCalendarEvent_Processed.txt
 
Description
-----------
This command will:
   1. Retrieve all users from Azure AD that have the Mail and EmployeeID attributes populated;
   2. Create a calendar event on their mailboxes, unless they are in the 'CreateCalendarEvent_Processed.txt' file.

.EXAMPLE
C:\PS> Get-Help .\CreateCalendarEvent.ps1 -Full
 
Description
-----------
Shows this help manual.
#>


[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $False, HelpMessage = "TXT file containing the email addresses of the mailboxes to create a calendar event on")]
    [ValidateScript({If ($_) {Test-Path $_}})]
    [String] $UsersFile,

    [Parameter(Position = 1, Mandatory = $False, HelpMessage = "TXT file containing the email addresses of the mailboxes NOT to create a calendar event on")]
    [ValidateScript({If ($_) {Test-Path $_}})]
    [String] $ExcludeUsersFile,

    [Parameter(Position = 2, Mandatory = $False, HelpMessage = "Creates a calendar event on all mailboxes of enabled users that have an EmployeeID configured")]
    [Switch] $AllUsers
)


#####################################################################################################
# Function to write all the actions performed by the script to a log file
#####################################################################################################
Function Write-Log {
    [CmdletBinding()]
    Param ([String] $Type, [String] $Message)
 
    $Logfile = $PSScriptRoot + "\CreateCalendarEvent_Log_$(Get-Date -f 'yyyyMMdd').txt"
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
# Script Start
#####################################################################################################
Write-Log -Type "INF" -Message "--------------------------------------------------------------------------"
Write-Log -Type "INF" -Message "START."

# Azure App Registration details
$ClientID = ""
$TenantID = ""
$CertThumprint = ""


# Check if the user ran the script with both the -Users parameter and the -All switch
If ($AllUsers -and $UsersFile) {
	Write-Log -Type "WRN" -Message "Please use only the -Users parameter or the -All switch, not both. Exiting script."
	Exit
}


#####################################################################################################
# Graph API
# Import Microsoft.Graph.Calendar module
Try {
	Import-Module Microsoft.Graph.Calendar -ErrorAction Stop
} Catch {
	Write-Log -Type "ERR" -Message "Unable to import 'Microsoft.Graph.Calendar' PowerShell Module: '$($_.Exception.Message)'. Please make sure you have it installed. Exiting script."
	Exit
}

# Import Microsoft.Graph.Users module, if required
If ($AllUsers) {
	Try {
		Import-Module Microsoft.Graph.Users -ErrorAction Stop
	} Catch {
		Write-Log -Type "ERR" -Message "Unable to import 'Microsoft.Graph.Users' PowerShell Module: '$($_.Exception.Message)'. Please make sure you have it installed. Exiting script."
		Exit
	}
}

# Connect to Graph API
Try {
	Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -CertificateThumbprint $CertThumprint -ErrorAction Stop
	# Welcome To Microsoft Graph!
	# Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
} Catch {
	Write-Log -Type "ERR" -Message "Unable to connect to MgGraph: '$($_.Exception.Message)'. Exiting script."
	Exit
}


#####################################################################################################
# Collect list of users to create a calendar for and those to exclude
# Import the file containing the users to NOT create a calender event for
If ($ExcludeUsersFile) {
	Try {
		$usersExclude = Get-Content $ExcludeUsersFile -ErrorAction Stop
	} Catch {
		Write-Log -Type "ERR" -Message "Unable to retrieve users to exclude from file '$ExcludeUsersFile': '$($_.Exception.Message)'. Exiting script."
		Exit
	}
}

# Import the file containing the users to create a calender event for
If ($UsersFile) {
	Try {
		$users = Get-Content $UsersFile -ErrorAction Stop
	} Catch {
		Write-Log -Type "ERR" -Message "Unable to retrieve users from file '$UsersFile': '$($_.Exception.Message)'. Exiting script."
		Exit
	}
}

# Get a list of ALL enabled users that have an EmployeeID to create a calender event for
If ($AllUsers) {
	Try {
		# Using UserPrincipalName instead of Mail to avoid recipients with the same SMTP proxy address (that might happen in Hybrid environments)
		$users = (Get-MgUser -All -Filter 'accountEnabled eq true' -Property AccountEnabled, EmployeeID, Mail, UserPrincipalName -ErrorAction Stop | ? {$_.EmployeeID -AND $_.EmployeeID -ne "EMPLOYEEID_EMPTY" -AND $_.Mail -AND $_.UserPrincipalName -notmatch "#EXT#@"}).UserPrincipalName
	} Catch {
		Write-Log -Type "ERR" -Message "Unable to retrieve all users: '$($_.Exception.Message)'. Exiting script."
		Exit
	}
}

# Gather some stats and initialise variables used for progress tracking
[Int] $count = 0
[Int] $countTotal = ($users | Measure).Count
If ($countTotal -eq 0) {
	Write-Log -Type "WRN" -Message "No users were retrieved."
	Disconnect-MgGraph
	Write-Log -Type "INF" -Message "END."
	Exit
} Else {Write-Log -Type "INF" -Message "Retrieved $("{0:N0}" -f $countTotal) user(s) to create a calendar event for"}


#####################################################################################################
# HTML code for the meeting invite
$eventBody = "<HTML>
	<body>
		<p class=MsoNormal><span style='font-family:""Arial"",sans-serif;color:#000000'>Whatever text you want.</span></p>
	</body>
</HTML>"


#####################################################################################################
# Process all users
ForEach ($user in $users) {
	# See if the user needs to be ignored/skipped
	If ($ExcludeUsersFile -AND $usersExclude -Contains $user) {
		Write-Log -Type "INF" -Message "Skipping '$user'"
		$count++
		Continue
	}

	# JSON for the meeting invite
	$params = @{
		Subject = "Event Subject"
		Body = @{
			ContentType = "HTML"
			Content = $eventBody
		}
		Start = @{
			DateTime = "2022-02-17T00:00:00"
			TimeZone = "GMT Standard Time"
		}
		End = @{
			DateTime = "2022-02-18T00:00:00"
			TimeZone = "GMT Standard Time"
		}
		Location = @{
			DisplayName = "See details below"
		}
		Attendees = @(
			@{
				EmailAddress = @{
					Address = $user
					Name = $user
				}
				Type = "required"
			}
		)
		AllowNewTimeProposals = $False
		HideAttendees = $True
		Importance = "Normal"
		IsAllDay = $True
		IsOnlineMeeting = $False
		IsReminderOn = $True
		ReminderMinutesBeforeStart = "15"
		ShowAs = "Free"
		TransactionId = $(New-Guid)
	}

	# Create the calendar invite on the user's mailbox
	Try {
		$reponse = New-MgUserEvent -UserId $user -BodyParameter $params -ErrorAction Stop
		Write-Log -Type "INF" -Message "Created Calendar event for '$user'"
		$user >> "CreateCalendarEvent_Processed.txt"
	} Catch {
		Write-Log -Type "ERR" -Message "Unable to write to '$user' calendar: $($_.Exception.Message)"
		$user >> "CreateCalendarEvent_Failed.txt"
	}

	$count++
	Write-Progress -Activity "Adding Calendar Events" -Status "Processed $("{0:N0}" -f $count) / $("{0:N0}" -f $countTotal)"
}

Disconnect-MgGraph
Write-Log -Type "INF" -Message "END."