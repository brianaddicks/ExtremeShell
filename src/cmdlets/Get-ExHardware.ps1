function Get-ExHardware {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowSupportOutput
	)
	
	$ScriptPrefix = "Get-ExHardware: "

    $NewDevice = New-Object -TypeName ExtremeShell.Device
	$Type      = Get-ExHardwareType $ShowSupportOutput
	
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
		
		switch ($Type) {
			'SecureStack' {
				# New Unit Start
				$Regex = [regex] "^\ +UNIT\ (?<num>\d+)\ HARDWARE\ INFORMATION"
				$Match = HelperEvalRegex $Regex $line
				if ($Match) {
					if (!($NewChassis)) {
						# Create NewChassis object on first Unit match
						$NewChassis        	 = New-Object -TypeName ExtremeShell.Chassis
						$NewChassis.Type     = "SecureStack"
						$NewDevice.Hardware += $NewChassis
						
						# Verbose
						$ChassisPrefix = $ScriptPrefix + $NewChassis.Type + ': '
						Write-Verbose $ChassisPrefix
					}
					$NewMember = New-Object -TypeName ExtremeShell.Slot
					$NewChassis.Slots += $NewMember
					$NewMember.Number = $Match.Groups['num'].Value
					
					# Verbose
					$MemberPrefix = "$ChassisPrefix Unit $($NewMember.Number): "
					Write-Verbose $MemberPrefix
				}
				
				if ($NewMember) {
					# Eval Parameters for this section
					$EvalParams = @{}
					$EvalParams.VariableToUpdate = ([REF]$NewMember)
					$EvalParams.ReturnGroupNum   = 1
					$EvalParams.LoopName         = 'fileloop'
					$EvalParams.StringToEval     = $line
					
					# Model
					$EvalParams.ObjectProperty = "Model"
					$EvalParams.Regex          = [regex] '^\ +Model:\ +(?<model>.+)'				
					$Eval                      = HelperEvalRegex @EvalParams
					
					# Serial
					$EvalParams.ObjectProperty = "SerialNumber"
					$EvalParams.Regex          = [regex] '^\ +Serial\ Number:\ +(?<serial>.+)'
					$Eval                      = HelperEvalRegex @EvalParams
					
					# Hardware Version
					$EvalParams.ObjectProperty = "HardwareVersion"
					$EvalParams.Regex          = [regex] '^\ +Hardware\ Version:\ +(?<version>.+)'
					$Eval                      = HelperEvalRegex @EvalParams
					
					# Firmware Version
					$EvalParams.ObjectProperty = "FirmwareVersion"
					$EvalParams.Regex          = [regex] '^\ +Firm[wW]are\ Version:\ +(?<version>.+)'
					$Eval                      = HelperEvalRegex @EvalParams
					
					# BootCode Version
					$EvalParams.ObjectProperty = "BootCodeVersion"
					$EvalParams.Regex          = [regex] '^\ +Boot\ Code\ Version:\ +(?<version>.+)'
					$Eval                      = HelperEvalRegex @EvalParams
				}
			}
			{ ($_ -eq 'S-Series') -or 
		      ($_ -eq 'K-Series') } {
				# New Chassis Start
				$Regex = [regex] "CHASSIS(\ (?<num>\d+))?\ HARDWARE\ INFORMATION"
				$Match = HelperEvalRegex $Regex $line
				if ($Match) {
					Write-Verbose "Line $i"
					# Create NewChassis object on first Unit match
					$NewChassis        	 = New-Object -TypeName ExtremeShell.Chassis
					$NewChassis.Number = $Match.Groups['num'].Value
					$NewDevice.Hardware += $NewChassis
					
					# Verbose
					$ChassisPrefix = $ScriptPrefix + 'Chassis ' + $NewChassis.Number + ': '
					Write-Verbose $ChassisPrefix
					continue
				}
				
				if ($NewChassis) {
					# Chassis Type
					$Regex = [regex] "Chassis\ Type:\ +(?<type>.+?)\("
					$Match = HelperEvalRegex $Regex $line
					if ($Match -and (!($NewChassis.Type))) {
						$NewChassis.Type = $Match.Groups['type'].Value.Trim()
						$NewChassis.PartNumber = (($NewChassis.Type -replace 'bonded','').Trim() -replace ' ','-').ToUpper()
						
						Write-Verbose "$ChassisPrefix $($NewChassis.Type)"
						continue
					}
					
					###########################################################################################
					# General Chassis Information
					
					# Eval Parameters for this section
					# Eval Parameters for this section
					$EvalParams = @{}
					$EvalParams.VariableToUpdate = ([REF]$NewChassis)
					$EvalParams.ReturnGroupNum   = 1
					$EvalParams.LoopName         = 'fileloop'
					$EvalParams.StringToEval     = $line
					
					# SerialNumber
					$EvalParams.ObjectProperty = "SerialNumber"
					$EvalParams.Regex          = [regex] "Chassis\ Serial\ Number:\ +(?<serial>.+)"
					$Eval                      = HelperEvalRegex @EvalParams
					
					# Version
					$EvalParams.ObjectProperty = "Version"
					$EvalParams.Regex          = [regex] "Chassis\ Version:\ +(?<version>.+)"
					$Eval                      = HelperEvalRegex @EvalParams
					
					# FanStatus
					$EvalParams.ObjectProperty = "FanStatus"
					$EvalParams.Regex          = [regex] "Chassis\ Fan:\ +(?<status>.+)"
					$Eval                      = HelperEvalRegex @EvalParams
					
					###########################################################################################
					# Power Supplies
					
					$Regex = [regex] "Chassis\ Power\ Supply\ (?<num>\d+):\ +(?<status>.+)"
					$Match = HelperEvalRegex $Regex $line
					if ($Match) {
						# Create NewPsu object on first Unit match
						$NewPsu = New-Object -TypeName ExtremeShell.PowerSupply
		                $NewPsu.Number = $Match.Groups['num'].Value
		                $NewPsu.Status = $Match.Groups['status'].Value
		                $NewChassis.PowerSupplies += $NewPsu
						
						# Verbose
						$PsuPrefix = $ChassisPrefix + " PSU " + $NewPsu.Number + ":"
						Write-Verbose "$PsuPrefix $($NewPsu.Status)"
						continue
					}
					
					if ($NewPsu) {
						# Eval Parameters for this section
						$EvalParams.VariableToUpdate = ([REF]$NewPsu)
						
						# PSU Type
						$EvalParams.ObjectProperty = "Type"
						$EvalParams.Regex          = [regex] "Type\ =\ (?<type>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
					}
					
					###########################################################################################
					# Slots
					
					# New Slot Start
					$Regex = [regex] "SLOT\ (?<num>\d+)(\ \(Chassis\ (?<chassis>\d+))?"
					$Match = HelperEvalRegex $Regex $line
					if ($Match) {
						$NewSlot        = New-Object -TypeName ExtremeShell.Slot
                        $NewSlot.Number = $Match.Groups['num'].Value
						
						$NewOptionModule = $null
						
						if ($Match.Groups['chassis'].Success) {
							$SlotChassis           = $Match.Groups['chassis'].Value
							$CorrectChassis        = $NewDevice.Hardware | ? { $_.Number -eq  $SlotChassis}
	                        $CorrectChassis.Slots += $NewSlot
						} else {
							$NewChassis.Slots += $NewSlot
						}

                        $SlotPrefix = ($ChassisPrefix -replace '\d+',$SlotChassis) + " Slot " + $NewSlot.Number + ":"
                        Write-Verbose $SlotPrefix
						continue
					}
					
					if ($NewSlot) {
						# Eval Parameters for this section
						if ($NewOptionModule) {
							$EvalParams.VariableToUpdate = ([REF]$NewOptionModule)
						} else {
							$EvalParams.VariableToUpdate = ([REF]$NewSlot)
						}
						
						# PSU Type
						$EvalParams.ObjectProperty = "Type"
						$EvalParams.Regex          = [regex] "Type\ =\ (?<type>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						
						# Model
						$EvalParams.ObjectProperty = "Model"
						$EvalParams.Regex          = [regex] "Model:\ +(?<model>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# SerialNumber
						$EvalParams.ObjectProperty = "SerialNumber"
						$EvalParams.Regex          = [regex] "^\ +Serial\ Number:\ +(?<serial>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# PartNumber
						$EvalParams.ObjectProperty = "PartNumber"
						$EvalParams.Regex          = [regex] "Part\ Number:\ +(?<num>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# HardwareVersion
						$EvalParams.ObjectProperty = "HardwareVersion"
						$EvalParams.Regex          = [regex] "Hardware\ Version:\ +(?<ver>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# FirmwareVersion
						$EvalParams.ObjectProperty = "FirmwareVersion"
						$EvalParams.Regex          = [regex] "Firmware\ Version:\ +(?<ver>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# BootCodeVersion
						$EvalParams.ObjectProperty = "BootCodeVersion"
						$EvalParams.Regex          = [regex] "BootCode\ Version:\ +(?<ver>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						# BootPromVersion
						$EvalParams.ObjectProperty = "BootPromVersion"
						$EvalParams.Regex          = [regex] "BootPROM\ Version:\ +(?<ver>.+)"
						$Eval                      = HelperEvalRegex @EvalParams
						
						###########################################################################################
						# OptionModules
						
						if ($Type -eq 'S-Series') {
							$Regex = [regex] "NIM\[(\d+)\]\ -\ Option"
							$Match = HelperEvalRegex $Regex $line
							if ($Match) {
								$NewOptionModule = New-Object -TypeName ExtremeShell.OptionModule
	                        	$NewSlot.OptionModules += $NewOptionModule
								$Nim = $Match.Groups[1].Value
	
								$OptionModulePrefix = $SlotPrefix + " Option Module " + $Nim + ":"
		                        Write-Verbose $OptionModulePrefix
								continue
							}
							
							if ($NewOptionModule) {
								# Eval Parameters for this section
								$EvalParams.VariableToUpdate = ([REF]$NewOptionModule)
								
								# Location
								$EvalParams.ObjectProperty = "Location"
								$EvalParams.Regex          = [regex] "(Chassis\ \d+)?\ +[Ll]ocation:\ +(?<loc>.+)"
								$EvalParams.ReturnGroupNum = 2
								$Eval                      = HelperEvalRegex @EvalParams
								$EvalParams.ReturnGroupNum = 1
								
								# BoardRevision
								$EvalParams.ObjectProperty = "BoardRevision"
								$EvalParams.Regex          = [regex] "Board\ Revision:\ +(?<rev>\d+)"
								$Eval                      = HelperEvalRegex @EvalParams
							}
						}
					}
				}
				
			}
			<#
			'K-Series' {
				Throw "K not handled yet"
			}#>
			default {
				Throw "Type not handled"
			}
		}
	}
	return $NewDevice
}
