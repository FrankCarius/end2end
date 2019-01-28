# Source  https://community.spiceworks.com/topic/2001221-powershell-script-to-upload-file-to-onedrive-business

# Specify tenant admin and site URL
#
#
$User = "justin.jacob@spidersoft.in"
$SiteURL = "https://test-my.sharepoint.com/personal/justin_jacob_spidersoftin";

$Folder = "C:\Users\justin.jacob\Desktop\New folder"
$DocLibName = "Documents"

#Add references to SharePoint client assemblies and authenticate to Office 365 site – required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$Password  = ConvertTo-SecureString ‘123@123’ -AsPlainText -Force

#Bind to site collection
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Password)

$Context.Credentials = $Creds

#Retrieve list
$List = $Context.Web.Lists.GetByTitle("$DocLibName")


$Context.Load($List)
$Context.ExecuteQuery()

#Upload file
Foreach ($File in (dir $Folder -File)) {
	$FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
	$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$FileCreationInfo.Overwrite = $true
	$FileCreationInfo.ContentStream = $FileStream
	$FileCreationInfo.URL = $File
	
	# $folder = $List.RootFolder.Folders | where {$_.Name -eq "TEST"} | select -First 1
	# $folder.Files.Add($FileCreationInfo)
	
	$Upload = $List.RootFolder.Files.Add($FileCreationInfo)
	$Context.Load($Upload)
	$Context.ExecuteQuery()
}