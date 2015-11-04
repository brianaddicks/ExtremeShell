function Get-ExPortStatus {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Parses "show neighbor" output.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowSupportOutput
	)
	
	Write-debug "wtf"
	$VerbosePrefix = "Get-ExPortStatus:"
	$ReturnObject  = @()
	
	$TotalLines = $ShowSupportOutput.Count
	$i          = 0 
	$StopWatch  = [System.Diagnostics.Stopwatch]::StartNew() # used by Write-Progress so it doesn't slow the whole function down
	
	Write-Verbose "testing"
	:fileloop foreach ($line in $ShowSupportOutput) {
		$i++
		$i
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
		
		$Regex = [regex] "^.+?->\ show\ port\ status"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) {
			$InSection = $true
			Write-Verbose "$VerbosePrefix In Section"
			continue
		}
		
		$Regex = [regex] "^Support->"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) { break }
		
		if ($InSection) {
			###########################################################################################
			# SpecialProperties
			$EvalParams = @{}
			$EvalParams.StringToEval   = $line
			
			# LagPorts
			$EvalParams.Regex = [regex] '(?mx)^
			                             (?<port>-+?)\ 
										 (?<alias>-+?)\ 
										 (?<oper>-+?)\ 
										 (?<admin>-+?)\ 
										 (?<speed>-+?)\ 
										 (?<duplex>-+?)\ 
										 (?<type>-+?)$'
											 
			$Eval             = HelperEvalRegex @EvalParams
			if ($Eval) {
				Write-Verbose "$VerbosePrefix Header matched"
				$Port = $Eval.Groups['port'].Value
				$Port.Length
			}
			<#
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
			#>
		}
	}	
	return $ReturnObject
}