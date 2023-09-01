# Documentation home: https://github.com/engrit-illinois/GpUpdate-Computer
function GpUpdate-Computer {

	param(
		[Parameter(Mandatory=$true,Position=0)]
		[string[]]$Queries,
		[string]$SearchBase="OU=Engineering,OU=Urbana,DC=ad,DC=uillinois,DC=edu",
		[PSCredential]$Credential,
		[int]$ThrottleLimit = 50
	)
	
	function log($msg) {
		Write-Host $msg
	}
	
	function Get-ADComputerName {
		param(
			[string[]]$Queries,
			[string]$SearchBase
		)
		Get-ADComputerLike -Queries $Queries -SearchBase $SearchBase | Select -ExpandProperty "Name"
	}

	function Get-ADComputerLike {
		param(
			[string[]]$Queries,
			[string]$SearchBase
		)
		$allResults = @()
		foreach($query in @($Queries)) {
			$results = Get-ADComputer -SearchBase $SearchBase -Filter "name -like `"$query`"" -Properties *
			$allResults += @($results)
		}
		$allResults
	}
	
	function Do-Stuff {
		$params = @{
			"Queries" = $Queries
		}
		if($SearchBase) { $params.SearchBase = $SearchBase }
		$comps = Get-AdComputerName @params
		
		log "Matching computers found:"
		$compsString = "`"" + ($comps -join "`",`"") + "`""
		log "    $compsString"
		
		# The following works on client computers with PS 5.1+
		$comps | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
			
			function log($msg, $noComp) {
				if(-not $noComp) { $msg = "[$_] $msg" }
				Write-Host $msg
			}
			
			log "Processing..."
			
			$scriptBlock = { echo "n" | gpupdate /force }
			$params = @{
				"ComputerName" = $_
				"ScriptBlock" = $scriptBlock
			}
			if($Credential) { $params.Credential = $Credential }
			
			$results = Invoke-Command @params
			
			log "Results:"
			# gpupdate returns an array of strings, with several blank lines
			$results | ForEach-Object {
				if($_) { log "    $_" $true }
			}
			log "Done processing."
		}
	}
	
	Do-Stuff
}