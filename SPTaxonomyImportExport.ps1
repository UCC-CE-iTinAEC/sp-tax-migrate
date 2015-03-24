function get-SPTagsFromPage {
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


function Export-SPTaxonomy {
	param(
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,
    [Parameter(Mandatory = $true)]
    [String]$OutputXmlPath
    )
	
	$webs = @()
	$web = Get-SPWeb -Identity $SiteUrl
	
	$myWeb = $web | Select Id, Title, ServerRelativeUrl
	$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $web)
	
	$webs += $myWeb
	
	$web.Webs | % {
		$myWeb = $_ | Select Id, Title, ServerRelativeUrl
		$myWeb | Add-Member -MemberType NoteProperty -Name "Pages" -Value (Get-SPPagesFromWeb -Web $_)
		$webs += $myWeb
	}
	
	Export-Clixml -Depth 9 -InputObject $webs -Path $OutputXmlPath
}

Export-SPTaxonomy -SiteUrl http://qa.oakgov.com/health -OutputXmlPath D:\health-tags.xml