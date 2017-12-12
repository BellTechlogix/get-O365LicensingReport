<#
 MSOL_Detailed_LicensingReport.ps1
 Adapted from Script Found Online
 Adapted By - TankCR
 Initial Date - 06Dec17
 Modified Date - 11Dec17
#>

#This Function creates a dialogue to return a Folder Path
function Get-Folder {
    param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

#File Select Function - Lets you select your input file
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File",
	  [Parameter(Mandatory=$False)]
      [string]$InitialDirectory
    )
	if(!($Title)) { $Title="Select Input File"}
	if(!($InitialDirectory)) { $InitialDirectory="c:\"}
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.FileName = $Title
	#can be set to filter file types
	IF($Filter -ne $null){
		$FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
		$OpenFileDialog.filter = $FilterString}
	if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
		$OpenFileDialog.filter = $Filter}
	$OpenFileDialog.ShowDialog() | Out-Null
	IF($OBJ){
		$fileobject = GI -Path $OpenFileDialog.FileName.tostring()
		Return $fileObject}
	else{Return $OpenFileDialog.FileName}
}

#This function allows you to decide on all users or some users
function Select-UserBase
{
	Param(
	[Parameter(Mandatory=$false)]
		[string] $selection
	)
	$title = "Select User Base"
	$message = "Do you wish to poll all O365 Accounts?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Selects All Mailboxes on Exchange."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Allows selection from import csv."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0)
    switch ($result)
    {
        0 {"Yes"}
        1 {"No"}
    }
    Return $selection
}


# Connect to Microsoft Online
if(!(get-module MSOnline)){Import-Module MSOnline}
$UserCredential = Get-Credential
Connect-MsolService -Credential $UserCredential

write-host "Connecting to Office 365..." -foregroundcolor green

#Date for timestamp
$date = Get-Date -Format yyyy_MM_dd-HHmm
# The Output will be written to this file in the selected Folder
$LogFile = ((Get-Folder)+"\"+$date+"_Office_365_Licenses.csv")

#select UserBase
$userbase = Select-UserBase
If(($userbase) -eq "No")
{
	write-host "Get Users From Input File" -ForegroundColor Green
	$MSOLUserFile = Get-FileName -Filter csv -Title "Select O365 Import File"  -Obj
	$MSOLUsers = Import-Csv $MSOLUserFile|foreach{get-msoluser -UserPrincipalName $_.emailaddress}
}

# Get a list of all licences that exist within the tenant
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1}

# Loop through all licence types found in the tenant
foreach ($license in $licensetype) 
{	
	# Build and write the Header for the CSV file
	$headerstring = "DisplayName,UserPrincipalName,AccountSku"
	
	foreach ($row in $($license.ServiceStatus)) 
	{
		$headerstring = ($headerstring + "," + $row.ServicePlan.servicename)
	}
	
	Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
	
	write-host ("Gathering users with the following subscription: " + $license.accountskuid)

	# Loop through all users and write them to the CSV file
	If(($userbase) -eq "Yes")
	{
		# Gather users for this particular AccountSku
    
		$users = Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses.accountskuid -contains $license.accountskuid}
		foreach ($user in $users) {
		
		write-host ("Processing " + $user.displayname)

        $thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid}

		$datastring = ($user.displayname + "," + $user.userprincipalname + "," + $license.SkuPartNumber)
		
		foreach ($row in $($thislicense.servicestatus)) {
			
			# Build data string
			$datastring = ($datastring + "," + $($row.provisioningstatus))
		}
		
		Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append
	}
	}
	ELSEIF(($userbase) -eq "No")
	{

		$users = $MSOLUsers| where {$_.isLicensed -eq "True" -and $_.licenses.accountskuid -contains $license.accountskuid}
		foreach ($user in $users) {
		
		write-host ("Processing " + $user.displayname)

        $thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid}

		$datastring = ($user.displayname + "," + $user.userprincipalname + "," + $license.SkuPartNumber)
		
		foreach ($row in $($thislicense.servicestatus)) {
			
			# Build data string
			$datastring = ($datastring + "," + $($row.provisioningstatus))
		}
		
		Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append
        $user = $null
	}
	}
    Out-File -FilePath $LogFile -InputObject " " -Encoding UTF8 -append
}			

write-host ("Script Completed.  Results available in " + $LogFile)
