function Get-ExLacp {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Parses "show neighbor" output.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowSupportOutput
	)
	
	$VerbosePrefix = "Get-ExLacp:"
	
	$TotalLines = $ShowSupportOutput.Count
	$i          = 0 
	$StopWatch  = [System.Diagnostics.Stopwatch]::StartNew() # used by Write-Progress so it doesn't slow the whole function down
	
	:fileloop foreach ($line in $ShowSupportOutput) {
		$i++
		
		# Write progress bar, we're only updating every 1000ms, if we do it every line it takes forever
		
		if ($StopWatch.Elapsed.TotalMilliseconds -ge 1000) {
			$PercentComplete = [math]::truncate($i / $TotalLines * 100)
	        Write-Progress -Activity "Reading Support Output" -Status "$PercentComplete% $i/$TotalLines" -PercentComplete $PercentComplete
	        $StopWatch.Reset()
			$StopWatch.Start()
		}
		
		if ($line -eq "") { continue }
		
		###########################################################################################
		# Check for the Start/Stop of lacp config
		
		$Regex = [regex] "^#\ lacp"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) {
			$LacpConfig = $true
			$NewObject = New-Object -Type ExtremeShell.LacpConfig
			continue
		}
		
		$Regex = [regex] "!"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) { $LacpConfig = $false }
		
		if ($LacpConfig) {
			###########################################################################################
			# SpecialProperties
			$EvalParams = @{}
			$EvalParams.StringToEval   = $line
			
			# LagPorts
			$EvalParams.Regex = [regex] '^set\ lacp\ aadminkey\ (?<name>\w+\.\d+\.\d+)\ (?<key>\d+)'
			$Eval             = HelperEvalRegex @EvalParams
			if ($Eval) {
				$NewLag                = New-Object -Type ExtremeShell.LagPort
				$NewLag.Name           = $Eval.Groups['name'].Value
				$NewLag.ActorAdminKey  = $Eval.Groups['key'].Value
				$NewObject.LagPorts   += $NewLag
			}
			
			# SinglePortLag
			$EvalParams.Regex = [regex] '^set\ lacp\ singleportlag\ enable'
			$Eval             = HelperEvalRegex @EvalParams
			if ($Eval) { $NewObject.SinglePortLag = $true }
			
			# FlowRegeneration
			$EvalParams.Regex = [regex] '^set\ lacp\ flowRegeneration\ enable'
			$Eval             = HelperEvalRegex @EvalParams
			if ($Eval) { $NewObject.FlowRegeneration = $true }
			
			###########################################################################################
			# Regular Properties
			$EvalParams.VariableToUpdate = ([REF]$NewObject)
			$EvalParams.ReturnGroupNum   = 1
			$EvalParams.LoopName         = 'fileloop'
				
			# SystemPriority
			$EvalParams.ObjectProperty = "SystemPriority"
			$EvalParams.Regex          = [regex] "^set\ lacp\ asyspri\ (\d+)"
			$Eval                      = HelperEvalRegex @EvalParams
			
			# OutportLocalPreference
			$EvalParams.ObjectProperty = "OutportLocalPreference"
			$EvalParams.Regex          = [regex] "^set\ lacp\ outportLocalPreference\ (.+)"
			$Eval                      = HelperEvalRegex @EvalParams
		}
		
		###########################################################################################
		# Check for the Start/Stop of lacp config
		
		$Regex = [regex] "^#\ port"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) {
			$PortConfig = $true
			continue
		}
		
		$Regex = [regex] "!"
		$Match = HelperEvalRegex $Regex $line
		if ($Match -and $PortConfig) { break }
		
		if ($PortConfig) {
			#set port lacp port tg.2.2 aadminkey 2100
			###########################################################################################
			# SpecialProperties
			$EvalParams = @{}
			$EvalParams.StringToEval   = $line
			
			# LagPorts
			$EvalParams.Regex = [regex] '^set\ port\ lacp\ port\ (?<member>\w+\.\d+\.\d+)\ aadminkey\ (?<key>\d+)'
			$Eval             = HelperEvalRegex @EvalParams
			if ($Eval) {
				$Lookup = $NewObject.LagPorts | ? { $_.ActorAdminKey -eq [int]($Eval.Groups['key'].Value) }
				$Lookup.MemberPorts += $Eval.Groups['member'].Value
			}
		}
	}	
	return $NewObject
}