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
	
	$VerbosePrefix = "Get-ExPortStatus:"
	$ReturnObject  = @()
	
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
		
		$Regex = [regex] "^.+?->\ show\ port\ status"
		$Match = HelperEvalRegex $Regex $line
		if ($Match) {
			$InSection = $true
			Write-Verbose "$VerbosePrefix In Section"
			continue
		}
		
		$Regex = [regex] "^Support->"
		$Match = HelperEvalRegex $Regex $line
		if ($Match -and $InSection) { break }
		
		if ($InSection) {
			###########################################################################################
			# SpecialProperties
			$EvalParams = @{}
			$EvalParams.StringToEval   = $line
			
			# Header
			$EvalParams.Regex = [regex] '(?mx)^
			                             (?<port>-+?)\ 
										 (?<alias>-+?)\ 
										 (?<oper>-+?)\ 
										 (?<admin>-+?)\ 
										 (?<speed>-+?)\ 
										 (?<duplex>-+?)\ 
										 (?<type>-+?)$'
											 
			$Eval = HelperEvalRegex @EvalParams
			if ($Eval) {
				Write-Verbose "$VerbosePrefix Header matched"
				$Port   = ($Eval.Groups['port'].Value).Length -1
				$Alias  = ($Eval.Groups['alias'].Value).Length
				$Oper   = ($Eval.Groups['oper'].Value).Length
				$Admin  = ($Eval.Groups['admin'].Value).Length
				$Speed  = ($Eval.Groups['speed'].Value).Length
				$Duplex = ($Eval.Groups['duplex'].Value).Length
				$Type   = ($Eval.Groups['type'].Value).Length
				
				$StatusRx  = "^(?<port>\w.{$Port})\ "
				$StatusRx += "(?<alias>.{$Alias})\ "
				$StatusRx += "(?<oper>.{$Oper})\ "
				$StatusRx += "(?<admin>.{$Admin})\ "
				$StatusRx += "(?<speed>.{$Speed})\ "
				$StatusRx += "(?<duplex>.{$Duplex})"
				$StatusRx += "(\ (?<type>.{$Type}))?"
				$Global:StatusRxTest = $StatusRx
				$StatusRx  = [regex]$StatusRx
				continue
			}
			
			# Status
			if ($StatusRx) {
				$EvalParams.Regex = $StatusRx
				$Eval             = HelperEvalRegex @EvalParams
				if ($Eval) {
					$NewObject     = New-Object -Type ExtremeShell.Port
					$ReturnObject += $NewObject
					
					$NewObject.Name       = $Eval.Groups['port'].Value.Trim()
					$NewObject.Alias      = $Eval.Groups['alias'].Value.Trim()
					$NewObject.OperStatus = $Eval.Groups['oper'].Value.Trim()
					$NewObject.Speed      = $Eval.Groups['speed'].Value.Trim()
					$NewObject.Duplex     = $Eval.Groups['duplex'].Value.Trim()
					$NewObject.Type       = $Eval.Groups['type'].Value.Trim() -replace "\ +","/"
					
					if ($Eval.Groups['admin'].Value.Trim() -eq "up") {
						$NewObject.Enabled = $true
					} else {
						$NewObject.Enabled = $false
					}
					continue
				}
				}
		}
	}	
	return $ReturnObject
}