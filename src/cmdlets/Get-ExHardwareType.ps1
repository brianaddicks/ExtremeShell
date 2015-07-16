function Get-ExHardwareType {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show support" from an Extreme Switch and determines the product family.  This is primarily for use with Get-ExHardware.
			
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowSupportOutput
	)
	
	# Determine type of device, currently supports:
	# B-Series
	# C-Series
	# S-Series
	# K-Series
	
	foreach ($line in $ShowSupportOutput) {
		if ($line -eq "") { continue }
		
		# Regex for chassis/member output start
		$SecureStackRx = [regex] "^\ +UNIT\ \d+\ HARDWARE\ INFORMATION"
		$CoreFlowRx    = [regex] "^CHASSIS\ (?<num>\d+\ )?HARDWARE\ INFORMATION"
		
		# Regex for K/S-Series
		$ChassisRx = [regex] "^\ +Chassis\ Type:\ +(?<bond>Bonded\ )?(?<type>S|K)\d+\ Chassis"
		
		
		# Check for Stackable output
		$Match = HelperEvalRegex -Regex $SecureStackRx -StringToEval $line
		if ($Match) {
			return "SecureStack"
		}
		
		# Check for CoreFlow output
		$Match = HelperEvalRegex -Regex $CoreFlowRx -StringToEval $line
		if ($Match) {
			$CoreFlow = $True
		}
		
		if ($CoreFlow) {
			$TypeMatch = HelperEvalRegex -Regex $ChassisRx -StringToEval $line
			if ($TypeMatch) {
				return "$($TypeMatch.Groups['type'].Value)-Series"
			}
		}
	}
	Throw "Switch not matched"
}