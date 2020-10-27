# //***************************************************************************
# // Author:	Jakob Gottlieb Svendsen, www.runbook.guru
# // Purpose:   Update a Sharepoint item using Tao Yangs SharePointSDK Module.
# // Get the module here: https://www.powershellgallery.com/packages/SharePointSDK/
# // USage Example:
# // Update Sharepoint
# //	.\Update-SPOnlineItemScript.ps1 `
# //			-ListName "Enable Office365 User" `
# //			-ListItemId $ItemID `
# //			-Values @{"Status" = "Completed"
# //						"Result" = "License Added"}
# //***************************************************************************
  			
Param(
	  	[Parameter(Mandatory=$true)][String]
		$ListName,
		[Parameter(Mandatory=$true)][int]
		$ListItemId,
		[Parameter(Mandatory=$true)][System.Collections.Hashtable]
		$Values
	)
	#$ListItemId = 22
	#$ListName = "DFS Share Request"
	#$Values = @{"Status" = "Complete"}
	$SPConnection = Get-AutomationConnection -Name 'SharePoint Online Connection'
		
	Write-Verbose "SharePoint Site URL: $($SPConnection.SharePointSiteURL)"
		
    #Get List Fields
    Import-Module SharePointSDK -ErrorAction "stop"
    $ListFields = Get-SPListFields -SPConnection $SPConnection -ListName $ListName -verbose  -ErrorAction "stop"
		
	$UpdateDetails = @{}
	foreach ($key in $Values.Keys)
	{
		$Field = ($ListFields | Where-Object {$_.Title -ieq $key -and $_.ReadOnlyField -eq $false}).InternalName
		if (!$Field) { throw "'$key' field/column not found in SharePoint List $ListName"}
		$UpdateDetails += @{ $Field = $Values[$key]}
	}
	$UpdateDetails
	
	Write-Verbose "`$UpdateDetails = $UpdateDetails"
	#Update a list item
    $UpdateListItem = Update-SPListItem -ListFieldsValues $UpdateDetails -ListItemID $ListItemID -ListName $ListName -SPConnection $SPConnection  -ErrorAction "stop"
    Write-Output "List Item (ID: $ListItemId) updated: $UpdateListItem."