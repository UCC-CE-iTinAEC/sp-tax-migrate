#region Import Methods
function PromptFor-File {
	param(
		[string]$Type = "Open",
		[string]$Title = "Select File",
		[string]$FileName = $null,
		[String[]]$FileTypes,
		[switch]$RestoreDirectory,
		[IO.DirectoryInfo]$InitialDirectory = $null
	)
	
	[Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	if($FileTypes) {
		$FileTypes | % {
			$filter += $_.ToUpper() + " Files|*.$_|"
		}
		$filter = $filter.TrimEnd("|")
	} else {
		$filter = "All Files|*.*"
	}
	
	switch($Type) {
		"Open" {
			$dialog = New-Object System.Windows.Forms.OpenFileDialog
			$dialog.Multiselect = $false
		}
		"Save" {
			$dialog = New-Object System.Windows.Forms.SaveFileDialog
		}
	}
	
	$dialog.FileName = $FileName
	$dialog.Title = $Title
	$dialog.Filter = $filter
	$dialog.RestoreDirectory = $RestoreDirectory
	$dialog.InitialDirectory = $InitialDirectory.Fullname
	$dialog.ShowHelp = $true
	if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
		return $dialog.FileName
	} else {
		return $null
	}
}

function Import-Terms {
	param(
		[Parameter(Mandatory = $true)]
		$Parent,
		[Parameter(Mandatory = $true)]
		[System.Collections.ArrayList]$Terms
	)
	$Terms | % {
		$Term = $_
		$txTerm = $Parent.CreateTerm($Term.Name, 1033, $Term.Id)
		$txTerm.IsAvailableForTagging = $Term.IsAvailableForTagging
		if ($Term.Terms.Count -gt 0) {
			Import-Terms -Parent $txTerm -Terms $Term.Terms
		}
	}
	
}

function Import-TermSet {
	param(
		[Parameter(Mandatory = $true)]
		[PSObject]$TermSet,
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPSite]$Site
	)
	
	$txSession = Get-SPTaxonomySession -Site $Site
	$txTermStore = $txSession.DefaultSiteCollectionTermStore
	$txSiteGroup = $txTermStore.GetSiteCollectionGroup($Site)
	$txTermSet = $txSiteGroup.CreateTermSet($TermSet.Name)
	Import-Terms -Parent $txTermSet -Terms $TermSet.Terms
	$txTermStore.CommitAll()
}

function Import-PageTags {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPFile]$File
	)
	if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($File.Web)) {
		$PubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($File.Web)
		if ([Microsoft.SharePoint.Publishing.PublishingPage]::IsPublishingPage($File.Item)) {
			$PubFile = [Microsoft.SharePoint.Publishing.PublishingPage]::GetPublishingPage($File.Item)
			$Item = $File.Item
#			if ($Item.
		}
	}
	
}

function Import-Web {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPSite]$Site,
		[Parameter(Mandatory = $true)]
		[PSObject]$Web
	)
	
	$spWeb = $Site.OpenWeb($Web.ServerRelativeUrl)
	
	$Web.Pages | % {
		$Page = $_
		if ($Page.Tags) {
			# process page
			$SPFile = $spWeb.GetFile($Page.Url)
		}
	}
	
	$spWeb.Dispose()
}

function Import-SPTaxonomy {
	param(
		[Parameter(Mandatory = $true)]
		[string]$SiteUrl,
		[Parameter(Mandatory = $false)]
		[string]$Path
	)
	
	if (-not $Path) {
	    # prompt user for location of xml file
		$Path = PromptFor-File -FileName "*.xml" -FileTypes xml -InitialDirectory $PWD.Path -Type Open -Title "Select XML File"
	}
	$objImport = Import-Clixml -Path $Path
	$Site = Get-SPSite -Identity $SiteUrl
	if ($objImport.TermSet) {
		Import-TermSet -TermSet $objImport.TermSet -Site $Site
	} else {
		Write-Error "No TermSet Found. Now exiting!"
		$Site.Dispose()
		return;
	}
	
	if ($objImport.Webs) {
		$objImport.Webs | % {
			$Web = $_
			Import-Web -Web $Web
		}
	} else {
		Write-Error "No Webs Found. Now exiting!"
		$Site.Dispose()
		return;
	}
	$Site.Dispose()
}

#endregion

#region Export Methods
function Get-SPTagsFromItem {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPListItem]$Item,
		[Parameter(Mandatory = $true)]
		[string]$FieldName
	)
#	$tags = New-Object -TypeName Microsoft.SharePoint.Taxonomy.TaxonomyFieldValueCollection -ArgumentList $PageItem["Tags"]
	
#	$tags | % {
#	Write-Host "`t$($_)"
#	}
	
	
	return $Item[$FieldName]
}

function Get-SPDocumentsFromWeb {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPWeb]$Web,
		[Parameter(Mandatory = $true)]
		[string]$FieldName
	)
	
	$docs = @()
	
	$dl = $Web.Lists.TryGetList("Documents")
	
	
	if ($dl) {
		$dl.Items | % {
			Write-Host "$($_.Name): $($_[$FieldName])"
			$myDoc = $_ | Select Id, Title, Url, Name
			$myDoc | Add-Member -MemberType NoteProperty -Name "Tags" -Value (Get-SPTagsFromItem -Item $_ -FieldName $FieldName)
			$docs += $myDoc
		}
	} else {
		$docs = $null
	}
	
	return $docs
}

function Get-SPPagesFromWeb {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPWeb]$Web,
		[Parameter(Mandatory = $false)]
		[string]$FieldName = "Tags"
	)
	$pages = @()
	
	$pl = $Web.Lists.TryGetList("Pages")
	if($pl) {
		$pl.Items | % {
			Write-Host "$($_.Name): $($_[$FieldName])"
			$myPage = $_ | Select Id, Title, Url, Name
			$myPage | Add-Member -MemberType NoteProperty -Name "Tags" -Value (Get-SPTagsFromItem -Item $_ -FieldName $FieldName)
#			$myPage | Add-Member -MemberType NoteProperty -Name "Tags" -Value $_["Tags"]
			$pages += $myPage
		}
	} else {
		$pages = $null
		
		
	}
	
	return $pages
}

function Export-SPTerms {
	param(
		[Parameter(Mandatory = $true)]
		$Terms
	)
	$txTerms = @()
	
	$Terms | % {
		$Term = $_
		$txTerm = $Term | Select Id, Name, IsAvailableForTagging
		if ($Term.TermsCount -gt 0) {
			$txTerm | Add-Member -MemberType NoteProperty -Name "Terms" -Value (Export-SPTerms -Terms $Term.Terms) 
		}
		$txTerms += $txTerm
	}
	return $txTerms
}

function Export-SPTermSet {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPSite]$Site,
		[Parameter(Mandatory = $false)]
		[string]$GroupName,
		[Parameter(Mandatory = $true)]
		[string]$TermSetName
	)
	
	$txSession = Get-SPTaxonomySession -Site $Site
	if ([string]::IsNullOrEmpty($GroupName)) {
		$txGroup = $txSession.DefaultSiteCollectionTermStore.GetSiteCollectionGroup($Site)
	} else {
		$txGroup = $txSession.DefaultSiteCollectionTermStore.Groups.Item($GroupName)
	}
	$txTermSet = $txGroup.TermSets.Item($TermSetName)
	
	$myTermSet = $txTermSet | Select Name
	$myTermSet | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $GroupName
	$myTermSet | Add-Member -MemberType NoteProperty -Name "Terms" -Value (Export-SPTerms -Terms $txTermSet.Terms)
		
	return $myTermSet
}

function Export-SPTaxonomy {
	param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,
    [Parameter(Mandatory = $true)]
    [string]$TermSetName,
    [Parameter(Mandatory = $true)]
    [string]$OutputXmlPath,
	[switch]$DocumentLibrary,
	[Parameter(Mandatory = $false)]
	[string]$GroupName,
	[Parameter(Mandatory = $false)]
	[string]$FieldName = "TaxKeyword"
    )
	
	$webs = @()
	$web = Get-SPWeb -Identity $SiteUrl
	
	$TermSet = Export-SPTermSet -Site $web.Site -GroupName $GroupName -TermSetName $TermSetName
	
	$myWeb = $web | Select Id, Title, ServerRelativeUrl
	if($DocumentLibrary) {
		$myWeb | Add-Member -MemberType NoteProperty -Name "Documents" -Value (Get-SPDocumentsFromWeb -Web $web -FieldName $FieldName)
	}
	else {
		$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $web)
	}
	
	$webs += $myWeb
	
	$web.Webs | % {
		$myWeb = $_ | Select Id, Title, ServerRelativeUrl
		if ($DocumentLibrary) {
			$myWeb | Add-Member -MemberType NoteProperty -Name "Documents" -Value (Get-SPDocumentsFromWeb -Web $web -FieldName $FieldName)
		} else {
			$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $_)
		}
		$webs += $myWeb
	}
	
	$exportObj = New-Object PSObject
	$exportObj | Add-Member -MemberType NoteProperty -Name "IsDocumentLibrary" -Value $DocumentLibrary
	$exportObj | Add-Member -MemberType NoteProperty -Name "TermSet" -Value $TermSet
	$exportObj | Add-Member -MemberType NoteProperty -Name "Webs" -Value $webs
	
	Export-Clixml -Depth 9 -InputObject $exportObj -Path $OutputXmlPath
}

#endregion

Export-SPTaxonomy -SiteUrl http://www.advantageoakland.com/ResearchPortal -GroupName "EDCA" -TermSetName "EDCA Terms" -DocumentLibrary -OutputXmlPath D:\researchportal-terms.xml
#Import-SPTaxonomy -SiteUrl http://six.sp-dev.us -Path E:\health-tags.xml