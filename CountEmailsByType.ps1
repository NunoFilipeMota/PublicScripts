Param (
      [Parameter(Position = 0, Mandatory = $False)]
      [String] $MailboxSMTP = "user@domain.com"
)

# Initialize the hash table that will hold the total number of emails per type
$hashClass = @{}
$hashSize = @{}
[Int] $countFolders = $countItems = $totalFolders = 0

#  Load the EWS Managed API DLL (do not forget to update the path if you install it on a different location/folder)
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[Void][Reflection.Assembly]::LoadFile($dllpath)

# Create a new object as an Exchange service and configure it to use AutoDiscover. Here I are using Exchange 2013, so update it to use the Exchange version you are targeting.
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)
$service.AutodiscoverUrl($MailboxSMTP)

# The script uses Impersonation to access the user’s mailbox. If you want to just use Full Access permissions, simply remove these following two lines.
$ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$MailboxSMTP
$service.ImpersonatedUserId = $ImpersonatedUserId

# Create another object and use a constructor to link a folder ID to the well-known folder MsgFolderRoot. If we wanted to process only the Inbox or Deleted Items, for example, we could use Inbox or DeletedItems
$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxSMTP)
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)

# Create a view, determine how many folders we will retrieve using this view, how we will traverse all folders and retrieve all folders based on these settings
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$ffResponse = $rfRootFolder.FindFolders($fvFolderView)

# Process each folder in the mailbox
$totalFolders = ($ffResponse.Folders).Count
ForEach ($ffFolder in $ffResponse.Folders) {
      # Create a new view, this time to determine how many emails we will retrieve
      $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)

      # Do/While statement to run as many passes on each folder as we need to
      Do{
            # Get the emails on the current folder
            $fiItems = $service.FindItems($ffFolder.Id,$ivItemView)

            # Process each email. Get the email class and update the hash table
            ForEach ($Item in $fiItems.Items) {
				$class = $Item.ItemClass
				$hashClass.Set_Item($class, $hashClass.Get_Item($class) + 1)

				$size = [Math]::Round($Item.Size/1024 , 0)
				#$size = $Item.Size
				$hashSize.Set_Item($class, $hashSize.Get_Item($class) + $size)
				
				$countItems++
				Write-Progress -Activity "Processing $($ffFolder.DisplayName)" -Status "Processed $countFolders/$totalFolders folders and $countItems items."
            }

            # Increment our offset by the number of emails we just processed
            $ivItemView.Offset += $fiItems.Items.Count
      } While ($fiItems.MoreAvailable)

      $countFolders++
}

# Print the final output
Write-Host "Processed $countFolders folders and $countItems items." -ForegroundColor Green
Write-Host "Items Class Count:" -ForegroundColor Yellow
$hashClass.GetEnumerator() | Sort Name | FT Name, Value -AutoSize

Write-Host "`n`nItems Class Size (KB):" -ForegroundColor Yellow
$hashSize.GetEnumerator() | Sort Name | FT Name, Value -AutoSize

#$hashSize.Keys | % {$newHashSize.Set_Item($_, [Math]::Round($hashSize.Get_Item($_) / 1024, 0))}
#$newHashSize.GetEnumerator() | Sort Name | FT Name, Value -AutoSize