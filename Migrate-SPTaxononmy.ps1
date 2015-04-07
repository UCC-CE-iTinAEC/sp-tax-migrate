# Migrate-SPTaxonomy

function Export-SPTaxonomy {
	param(
	[Parameter(Mandatory = $true)]
	[Uri]$SiteUrl,
	$Groups)
	
	$site = Get-SPSite -Identity $SiteUrl
	
	$txSession = Get-SPTaxonomySession -Site $site
	$TermStore = $txSession.DefaultSiteCollectionTermStore
	
	$oTermGroups = @()
	
	$Groups | % {
		$GroupName = $_
		$TermGroup = $TermStore.Groups | ? { $_.Name -eq $GroupName }
		$TermGroup.Name
		$TermGroup.TermSets | % {
			$_.Name
		}
	}
	
	$site.Dispose()
}

Export-SPTaxonomy -SiteUrl "http://www.advantageoakland.com" -Groups "EDCA"