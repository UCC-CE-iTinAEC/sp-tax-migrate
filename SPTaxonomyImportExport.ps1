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

function Import-TermSet {
	param(
		[Parameter(Mandatory = $true)]
		[PSObject]$TermSet,
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPSite]$Site
	)
	
	$txSession = Get-SPTaxonomySession -Site $Site
	$txSiteGroup = $txSession.DefaultSiteCollectionTermStore.GetSiteCollectionGroup($Site)
	$txTermSet = $txSiteGroup.CreateTermSet($TermSet.Name)
	
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
		Import-TermSet -TermSet $objImport.TermSet
	} else {
		Write-Error "No TermSet Found. Now exiting!"
		$Site.Dispose()
		return;
	}
	
	if ($objImport.Webs) {
	} else {
		Write-Error "No Webs Found. Now exiting!"
		$Site.Dispose()
		return;
	}
	$Site.Dispose()
}

#endregion

#region Export Methods
function Get-SPTagsFromPage {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPListItem]$PageItem
	)
#	$tags = New-Object -TypeName Microsoft.SharePoint.Taxonomy.TaxonomyFieldValueCollection -ArgumentList $PageItem["Tags"]
	
#	$tags | % {
#	Write-Host "`t$($_)"
#	}
	
	
	return $PageItem["Tags"]
}

function Get-SPPagesFromWeb {
	param(
		[Parameter(Mandatory = $true)]
		[Microsoft.SharePoint.SPWeb]$Web
	)
	$pages = @()
	
	$pl = $Web.Lists.TryGetList("Pages")
	if($pl) {
		$pl.Items | % {
			Write-Host "$($_.Name): $($_['Tags'])"
			$myPage = $_ | Select Id, Title, Url, Name
			$myPage | Add-Member -MemberType NoteProperty -Name "Tags" -Value (Get-SPTagsFromPage -PageItem $_)
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
		[Microsoft.SharePoint.Taxonomy.Generic.TaxonomyItemCollection`1[Microsoft.SharePoint.Taxonomy.Term]]$Terms
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
		[Parameter(Mandatory = $true)]
		[string]$TermSetName
	)
	
	$txSession = Get-SPTaxonomySession -Site $Site
	$txSiteGroup = $txSession.DefaultSiteCollectionTermStore.GetSiteCollectionGroup($Site)
	$txTermSet = $txSiteGroup.TermSets.Item($TermSetName)
	
	$myTermSet = $txTermSet | Select Name
	$myTermSet | Add-Member -MemberType NoteProperty -Name "Terms" -Value (Export-SPTerms -Terms $txTermSet.Terms)
	
	return $myTermSet
}

function Export-SPTaxonomy {
	param(
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,
    [Parameter(Mandatory = $true)]
    [String]$TermSetName,
    [Parameter(Mandatory = $true)]
    [String]$OutputXmlPath
    )
	
	$webs = @()
	$web = Get-SPWeb -Identity $SiteUrl
	
	$TermSet = Export-SPTermSet -Site $web.Site -TermSetName $TermSetName
	
	$myWeb = $web | Select Id, Title, ServerRelativeUrl
	$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $web)
	
	$webs += $myWeb
	
	$web.Webs | % {
		$myWeb = $_ | Select Id, Title, ServerRelativeUrl
		$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $_)
		$webs += $myWeb
	}
	
	$exportObj = New-Object PSObject
	$exportObj | Add-Member -MemberType NoteProperty -Name "TermSet" -Value $TermSet
	$exportObj | Add-Member -MemberType NoteProperty -Name "Webs" -Value $webs
	
	Export-Clixml -Depth 9 -InputObject $exportObj -Path $OutputXmlPath
}

#endregion

# Export-SPTaxonomy -SiteUrl http://qa.oakgov.com/health -TermSetName "Health Tags" -OutputXmlPath D:\health-tags.xml