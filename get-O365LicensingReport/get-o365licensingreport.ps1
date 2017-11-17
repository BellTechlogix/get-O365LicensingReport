#
# get-o365licensingscript.ps1
#

import-module MSOnline
$O365Cred = Get-Credential
Connect-MsolService -Credential $O365Cred

#Function to select a file
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.FileName = $Title
  #can be set to filter file types
  IF($Filter -ne $null){
  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
	$OpenFileDialog.filter = $FilterString}
  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
  $OpenFileDialog.filter = $Filter
  }
  $OpenFileDialog.ShowDialog() | Out-Null
  ## dont bother asking, just give back the object
  IF($OBJ){
  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
  Return $fileObject
  }
  else{Return $OpenFileDialog.FileName}
}

<#
	This Function creates a dialogue to return a Folder Path
#>
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


#Select your list from a csv
#Uses get-filename function to select userlist and import it
$file = get-filename -Filter "csv" -Title "Select User List" -obj
$userlist = import-csv $file.FullName

#wipes any previous $output data
$output = $null

#Builds a blank array to store the data
$output=@()

#Goes through each User in the imported Userlist to build the report
ForEach($user in $userlist)
{
    #Gets the license status as assigned to each user
	$userlic = (Get-MsolUser -UserPrincipalName $user.EmailAddress).Licenses.ServiceStatus
    #start building the individual objects to go into the array
	$userObj = New-Object PSObject
	#Creates Property with email address
    $userObj | Add-Member NoteProperty -Name "EmailAddress" -Value $user.EmailAddress
	#Creates Property with EXCHANGE_S_ENTERPRISE license data
    $userObj | Add-Member NoteProperty -Name "Exchange Enterprise" -Value ($userlic|where{$_.ServicePlan.ServiceName -eq 'EXCHANGE_S_ENTERPRISE'}).ProvisioningStatus
	#Creates Property with RMS_S_PREMIUM license data
    $userObj | Add-Member NoteProperty -Name "Azure Rights Management Premium" -Value ($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_PREMIUM'}).ProvisioningStatus
	#Creates Property with RMS_S_Enterprise license data - Checks if it has more then one value and loops it to combine them into a string
    If((($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE'}).ProvisioningStatus).count -gt 1)
    {
        #Gets first value in the multi value object
		$val01 = ((($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE'}).ProvisioningStatus)[0])
        #Gets second value in the multi value object
        $val02 = ((($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE'}).ProvisioningStatus)[1])
        #combines the values to one string
        $combine = "$val01,$val02"
        #adds the combined values to the property
        $userObj | Add-Member NoteProperty -Name "Azure Rights Management Enterprise" -Value $combine
    }
    #Checks if RMS_S_ENTERPRISE is a value less than 2 then adds the property vvalue if it is
	If((($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE'}).ProvisioningStatus).count -lt 2)    
    {$userObj | Add-Member NoteProperty -Name "Azure Rights Management Enterprise" -Value (($userlic|where{$_.ServicePlan.ServiceName -eq 'RMS_S_ENTERPRISE'}).ProvisioningStatus|out-string)}
 	#Creates Property with AAD_PREMIUM license data   
	$userObj | Add-Member NoteProperty -Name "Azure Active Directory Premium" -Value ($userlic|where{$_.ServicePlan.ServiceName -eq 'AAD_PREMIUM'}).ProvisioningStatus
    #Adds the combined values to the $Output Array
	$output += $userObj
	#Nulls out the userobject to ensure next loop doesn't grab data from the last loop
    $userObj = $null
}
#runs the get-folder function to grab location for output then names the output file and exports it 
$output|export-csv ((Get-Folder)+"\O365Licensing_"+($file.Name)) -NoTypeInformation
