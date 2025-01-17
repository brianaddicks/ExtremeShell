###############################################################################
## Start Powershell Cmdlets
###############################################################################

###############################################################################
# Get-EntHardware

function Get-EntHardware {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowHardwareOutput
	)

        
    $ScriptPrefix = "Get-EntHardware: "

    $NewDevice = New-Object -TypeName ExtremeShell.Device

    foreach ($c in $ShowHardwareOutput) {

        $ChassisStartRx   = [regex] "CHASSIS(\ (?<num>\d+))?\ HARDWARE\ INFORMATION"
        $ChassisStartMatch = $ChassisStartRx.Match($c)
        if ($ChassisStartMatch.Success) {
            $Chassis = $true
            $NewChassis = New-Object -TypeName ExtremeShell.Chassis
            $NewChassis.Number = $ChassisStartMatch.Groups['num'].Value
            $NewDevice.Hardware += $NewChassis

            $ChassisPrefix = $ScriptPrefix + "Chassis " + $NewChassis.Number + ":"
            Write-Verbose $ChassisPrefix
        }

        if ($Chassis) {
            $ChassisTypeRx    = [regex] "Chassis\ Type:\ +(?<type>.+?)\("
            $ChassisTypeMatch = $ChassisTypeRx.Match($c)
            if ($ChassisTypeMatch.Success) {
                $NewChassis.Type = $ChassisTypeMatch.Groups['type'].Value.Trim()

                switch ($NewChassis.Type) {
                    { ($_ -match "bonded") -or
                      ($_ -match "K10") } {
                        $NewChassis.PartNumber = (($_ -replace "bonded","").Trim() -replace " ","-").ToUpper()
                    }
                    default {
                        Throw "$_ not handled"
                    }
                }

                Write-Verbose "$ChassisPrefix $($NewChassis.Type)"
            }

            $ChassisSerialRx  = [regex] "Chassis\ Serial\ Number:\ +(?<serial>.+)"
            $ChassisSerialMatch = $ChassisSerialRx.Match($c)
            if ($ChassisSerialMatch.Success) {
                $NewChassis.SerialNumber = $ChassisSerialMatch.Groups['serial'].Value

                Write-Verbose "$ChassisPrefix $($NewChassis.SerialNumber)"
            }

            $ChassisVersionRx = [regex] "Chassis\ Version:\ +(?<version>.+)"
            $ChassisVersionMatch = $ChassisVersionRx.Match($c)
            if ($ChassisVersionMatch.Success) {
                $NewChassis.Version = $ChassisVersionMatch.Groups['version'].Value

                Write-Verbose "$ChassisPrefix $($NewChassis.Version)"
            }

            $ChassisFanRx     = [regex] "Chassis\ Fan:\ +(?<status>.+)"
            $ChassisFanMatch = $ChassisFanRx.Match($c)
            if ($ChassisFanMatch.Success) {
                $NewChassis.FanStatus = $ChassisFanMatch.Groups['status'].Value

                Write-Verbose "$ChassisPrefix $($NewChassis.FanStatus)"
            }

            ################################################################################
            # Psu
    
            $PsuStartRx = [regex] "Chassis\ Power\ Supply\ (?<num>\d+):\ +(?<status>.+)"
            $PsuStartMatch = $PsuStartRx.Match($c)
            if ($PsuStartMatch.Success) {
                $NewPsu = New-Object -TypeName ExtremeShell.PowerSupply
                $NewPsu.Number = $PsuStartMatch.Groups['num'].Value
                $NewPsu.Status = $PsuStartMatch.Groups['status'].Value
                $NewChassis.PowerSupplies += $NewPsu

                $PsuPrefix = $ChassisPrefix + " PSU " + $NewPsu.Number + ":"
                Write-Verbose "$PsuPrefix $($NewPsu.Status)"
            }

            $PsuTypeRx  = [regex] "Type\ =\ (?<type>.+)"
            $PsuTypeMatch = $PsuTypeRx.Match($c)
            if ($PsuTypeMatch.Success) {
                $NewPsu.Type = $PsuTypeMatch.Groups['type'].Value
            
                Write-Verbose "$PsuPrefix $($NewPsu.Type)"
            }

            ################################################################################
            # Slots

            switch ($NewChassis.PartNumber) {
                "S4-CHASSIS" {
                    $SlotStartRx = [regex] "SLOT\ (?<num>\d+)\ \(Chassis\ (?<chassis>\d+)"
                    $SlotStartMatch = $SlotStartRx.Match($c)
                    if ($SlotStartMatch.Success) {
                        $OptionModule = $false
                        $Slot = $true
                        $NewItem = New-Object -TypeName ExtremeShell.Slot
                        $NewItem.Number = $SlotStartMatch.Groups['num'].Value
                        $SlotChassis = $SlotStartMatch.Groups['chassis'].Value
            
                        $CorrectChassis = $NewDevice.Hardware | ? { $_.Number -eq $SlotChassis }
                        $CorrectChassis.Slots += $NewItem

                        $SlotPrefix = ($ChassisPrefix -replace '\d+',$SlotChassis) + " Slot " + $NewItem.Number + ":"
                        Write-Verbose "$SlotPrefix $SlotChassis"
                    }

                    $SlotModelRx = [regex] "Model:\ +(?<model>.+)"
                    $SlotModelMatch = $SlotModelRx.Match($c)
                    if ($SlotModelMatch.Success) {
                        if ($OptionModule) {
                            if (!($NewOptionModule.Model)) {
                                $NewOptionModule.Model = $SlotModelMatch.Groups['model'].Value
                                Write-Verbose "$OptionPrefix $($NewOptionModule.Model)"
                            }
                        } else {
                            if (!($NewItem.Model)) {
                                $NewItem.Model = $SlotModelMatch.Groups['model'].Value
                                Write-Verbose "$SlotPrefix $($NewItem.Model)"
                            }
                        }
                    }

                    $SerialNumberRx = [regex] "^\ +Serial\ Number:\ +(?<serial>.+)"
                    $SerialNumberMatch = $SerialNumberRx.Match($c)
                    if ($SerialNumberMatch.Success) {
                        if ($OptionModule) {
                            if (!($NewOptionModule.SerialNumber)) {
                                $NewOptionModule.SerialNumber = $SerialNumberMatch.Groups['serial'].Value
                                Write-Verbose "$OptionPrefix $($NewOptionModule.SerialNumber)"
                            }
                        } elseif ($Slot) {
                            if (!($NewItem.SerialNumber)) {
                                $NewItem.SerialNumber = $SerialNumberMatch.Groups['serial'].Value
                                Write-Verbose "$SlotPrefix $($NewItem.SerialNumber)"
                            }
                        }
                    }

                    $PartNumberRx = [regex] "Part\ Number:\ +(?<num>.+)"
                    $PartNumberMatch = $PartNumberRx.Match($c)
                    if ($PartNumberMatch.Success) {
                        if ($OptionModule) {
                            $NewOptionModule.PartNumber = $PartNumberMatch.Groups['num'].Value
                            Write-Verbose "$OptionPrefix $($NewOptionModule.PartNumber)"
                        } else {
                            $NewItem.PartNumber = $PartNumberMatch.Groups['num'].Value
                            Write-Verbose "$SlotPrefix $($NewItem.PartNumber)"
                        }
                    }

                    $HardwareVersionRx = [regex] "Hardware\ Version:\ +(?<ver>.+)"
                    $HardwareVersionMatch = $HardwareVersionRx.Match($c)
                    if ($HardwareVersionMatch.Success) {
                        $NewItem.HardwareVersion = $HardwareVersionMatch.Groups['ver'].Value
                        Write-Verbose "$SlotPrefix $($NewItem.HardwareVersion)"
                    }

                    $FirmwareVersionRx = [regex] "Firmware\ Version:\ +(?<ver>.+)"
                    $FirmwareVersionMatch = $FirmwareVersionRx.Match($c)
                    if ($FirmwareVersionMatch.Success) {
                        $NewItem.FirmwareVersion = $FirmwareVersionMatch.Groups['ver'].Value
                        Write-Verbose "$SlotPrefix $($NewItem.FirmwareVersion)"
                    }

                    $BootCodeVersionRx = [regex] "BootCode\ Version:\ +(?<ver>.+)"
                    $BootCodeVersionMatch = $BootCodeVersionRx.Match($c)
                    if ($BootCodeVersionMatch.Success) {
                        $NewItem.BootCodeVersion = $BootCodeVersionMatch.Groups['ver'].Value
                        Write-Verbose "$SlotPrefix $($NewItem.BootCodeVersion)"
                    }

                    $BootPROMVersionRx = [regex] "BootPROM\ Version:\ +(?<ver>.+)"
                    $BootPROMVersionMatch = $BootPROMVersionRx.Match($c)
                    if ($BootPROMVersionMatch.Success) {
                        $NewItem.BootPromVersion = $BootPROMVersionMatch.Groups['ver'].Value
                        Write-Verbose "$SlotPrefix $($NewItem.BootPromVersion)"
                    }

                    ################################################################################
                    # Option Module

                    $OptionModuleRx = [regex] "NIM\[(\d+)\]\ -\ Option"
                    $OptionModuleMatch = $OptionModuleRx.Match($c)
                    if ($OptionModuleMatch.Success) {
                        $OptionModule = $true
                        $NewOptionModule = New-Object -TypeName ExtremeShell.OptionModule
                        $NewItem.OptionModules += $NewOptionModule
                        $Nim = $OptionModuleMatch.Groups[1].Value

                        $OptionPrefix = $SlotPrefix + " Option Module " + $Nim + ":"
                        Write-Verbose $OptionPrefix
                    }

                    $OptionModuleLocRx = [regex] "Chassis\ \d+\ location:\ +(?<loc>.+)"
                    $OptionModuleLocMatch = $OptionModuleLocRx.Match($c)
                    if ($OptionModuleLocMatch.Success) {
                        $NewOptionModule.Location = $OptionModuleLocMatch.Groups['loc'].Value
                        Write-Verbose "$OptionPrefix $($NewOptionModule.Location)"
                    }

                    $OptionModuleRevRx = [regex] "Board\ Revision:\ +(?<rev>\d+)"
                    $OptionModuleRevMatch = $OptionModuleRevRx.Match($c)
                    if ($OptionModuleRevMatch.Success) {
                        $NewOptionModule.BoardRevision = $OptionModuleRevMatch.Groups['rev'].Value
                        Write-Verbose "$OptionPrefix $($NewOptionModule.BoardRevision)"
                    }
                }
                "K10-CHASSIS" {
                    $SlotStartRx = [regex] "SLOT\ (?<num>\d+)"
                    $SlotStartMatch = $SlotStartRx.Match($c)
                    if ($SlotStartMatch.Success) {
                        $OptionModule = $false
                        $Slot = $true
                        $NewItem = New-Object -TypeName ExtremeShell.Slot
                        $NewItem.Number = $SlotStartMatch.Groups['num'].Value
                        $SlotChassis = $SlotStartMatch.Groups['chassis'].Value
            
                        $CorrectChassis = $NewDevice.Hardware | ? { $_.Number -eq $SlotChassis }
                        $CorrectChassis.Slots += $NewItem

                        $SlotPrefix = ($ChassisPrefix -replace '\d+',$SlotChassis) + " Slot " + $NewItem.Number + ":"
                        Write-Verbose "$SlotPrefix $SlotChassis"
                    }

                    $SlotModelRx = [regex] "Model:\ +(?<model>.+)"
                    $SlotModelMatch = $SlotModelRx.Match($c)
                    if ($SlotModelMatch.Success) {
                        if (!($NewItem.Model)) {
                            $NewItem.Model = $SlotModelMatch.Groups['model'].Value
                            Write-Verbose "$SlotPrefix $($NewItem.Model)"
                        }
                    }

                    $SerialNumberRx = [regex] "^\s+Serial\ Number:\ +(?<serial>.+)"
                    $SerialNumberMatch = $SerialNumberRx.Match($c)
                    if ($SerialNumberMatch.Success) {
                        if (!($NewItem.SerialNumber)) {
                            $NewItem.SerialNumber = $SerialNumberMatch.Groups['serial'].Value
                            Write-Verbose "$SlotPrefix $($NewItem.SerialNumber)"
                        }
                    }

                    $PartNumberRx = [regex] "Part\ Number:\ +(?<num>.+)"
                    $PartNumberMatch = $PartNumberRx.Match($c)
                    if ($PartNumberMatch.Success) {
                        $NewItem.PartNumber = $PartNumberMatch.Groups['num'].Value
                        Write-Verbose "$SlotPrefix $($NewItem.PartNumber)"
                    }
                }
            }
        }
    }

    return $NewDevice
}

###############################################################################
# Get-ExHardware

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

###############################################################################
# Get-ExHardwareType

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

###############################################################################
# Get-ExLacp

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

###############################################################################
# Get-ExNeighbor

function Get-ExNeighbor {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Parses "show neighbor" output.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[array]$ShowSupportOutput
	)
	
	$VerbosePrefix = "Get-ExNeighbor:"
	
	$IpRx         = [regex] "(\d+\.){3}\d+"
	$PromptString = [regex] "^.+?->"
	$StartString  = [regex] "$PromptString\ show\ neighbors"
	
	$TotalLines = $ShowSupportOutput.Count
	$i          = 0 
	$StopWatch  = [System.Diagnostics.Stopwatch]::StartNew() # used by Write-Progress so it doesn't slow the whole function down
	
	$ReturnObject = @()
	
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
		# Check for the Start/Stop
		
		$Regex = $StartString
		$Match = HelperEvalRegex $Regex $line
		if ($Match) {
			$InSection = $true
			continue
		}
		
		$Regex = $PromptString
		$Match = HelperEvalRegex $Regex $line
		if ($Match) { $InSection = $false }
		
		if ($InSection) {
			$Regex = [regex] "^(?<localport>\w+\.\d+\.\d+)\ +(?<deviceid>[^\ ]+?)\ +(?<remoteport>[^\ ]+?)\ +(?<type>\w+)\ +(?<ip>$IpRx)?"
			$Match = HelperEvalRegex $Regex $line
			if ($Match) {
				$NewObject             = New-Object -Type ExtremeShell.Neighbor
				$NewObject.LocalPort   = $Match.Groups['localport'].Value
				$NewObject.DeviceId    = $Match.Groups['deviceid'].Value
				$NewObject.RemotePort  = $Match.Groups['remoteport'].Value
				$NewObject.Type        = $Match.Groups['type'].Value
				$NewObject.IpAddress   = $Match.Groups['ip'].Value
				$ReturnObject         += $NewObject
			}
		}
	}	
	return $ReturnObject
}

###############################################################################
# Get-ExPortStatus

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

###############################################################################
# Get-ExSModuleClass

function Get-ExSModuleClass {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Decodes S-Series Modules to their appropriate Class designation.  Information based on https://community.extremenetworks.com/extreme/topics/s_series_module_decoder-1i4xc8			
			
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[string]$ModulePartNumber
	)
	
	$SModuleRx = [regex] "(?x)
		 				  (?<devicetype>S|SO|SSA-)
                          (?<interfacetype>G|K|L|T|V)
                          (?<class>\d)
                          (?<option>\d)
                          (?<porttype>\d{2})
                          -
                          (?<throughput>\d{2})
                          (?<portcount>\d{2})
                          (
                          -
                          (?<fabriccapacity>F\d)
                          )?

					     "

    $Match = $SModuleRx.Match($ModulePartNumber)
    
    $DeviceTypes = @{ 'S'   = 'S-Series Chassis'
                      'SO'  = 'S-Series Option/Expansion Module'
                      'SSA' = 'S-Series Standalone' }
    
    $InterfaceTypes = @{ 'G' = '1Gb SFP/SFP+'
                         'K' = '10Gb'
                         'L' = '40Gb'
                         'T' = '10/100/1000 BaseT'
                         'V' = 'VSB' }

    $Classes = @{ '4' = 'S130'
                  '2' = 'S140'
                  '1' = 'S150'
                  '5' = 'S155'
                  '8' = 'S180' }

    $OptionModuleSeries = @{ '1' = 'Type1'
                             '2' = 'Type2'
                             '3' = 'S140/180 Expansion Module' }

    $OptionModuleType = @{ '0' = 'Expansion Module'
                           '2' = 'Standard' }

    $PortType = @{ '01' = 'SFP'
                   '06' = '10/100/1000Mb RJ45 POE'
                   '08' = 'SFP+'
                   '09' = '10Gb and VSB'
                   '13' = '40Gb and VSB'
                   '18' = 'SFP/SFP+'
                   '28' = '10/100/1000Mb RJ45 and 10Gb SFP+'
                   '68' = '10/100/1000Mb RJ45 POE and 10Gb SFP+' }

    $IoCapacity = @{ '01' = '10Gb'
                     '02' = '20Gb'
                     '03' = '30Gb'
                     '06' = '60Gb'
                     '08' = '80Gb'
                     '12' = '160Gb' }

    $FabricCapacity = @{ 'F6' = '1280Gb'
                         'F8' = '2560Gb' }

    
    $DeviceType = $DeviceTypes.Get_Item($Match.Groups['devicetype'].Value)
    "DeviceType: $DeviceType"

    $InterfaceType = $InterfaceTypes.Get_Item($Match.Groups['interfacetype'].Value)
    "InterfaceType: $InterfaceType"

    switch ($DeviceType) {
        { ($_ -eq 'S-Series Chassis') -or
          ($_ -eq 'S-Series Standalone') } {
            $Class = $Classes.Get_Item($Match.Groups['class'].Value)
        }
        'S-Series Option/Expansion Module' {
            $Class = $OptionModuleSeries.Get_Item($Match.Groups['class'].Value)
        }
    }
    "Class: $Class"
}

Get-ExSModuleClass 'SG4101-0248'


###############################################################################
# New-ExVisioK10

function New-ExVisioK10 {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[ExtremeShell.Device]$DeviceObject,
		
		[Parameter(Mandatory=$False)]
		[string]$VsdPath,
		
		[Parameter(Mandatory=$False)]
		[string]$SvgPath,
		
		[Parameter(Mandatory=$False)]
		[string]$PngPath
	)
	
	# Used for Write-Verbose and Throw
	$VerbosePrefix = "New-ExVisioK6:"
	
	# Test for VisioShell
	$IsModuleAvailable = Get-Module -Listavailable -Name VisioShell
	if (!($IsModuleAvailable)) {
		Throw "$VerbosePrefex VisioShell Module not found."
	} else {
		$Import = Import-Module VisioShell
	}
	
	# Create objects for blade/psu placement
	$BladeCoords = @()
	$BladeCoords += HelperBladeProps slot1 0.25 9.738281 2.0625 9.363281 3.3406 9.5503
	$BladeCoords += HelperBladeProps slot2 7.09375 9.738281 8.90625 9.363281 5.4676 9.5503
	$BladeCoords += HelperBladeProps slot3 0.25 9.316406 2.0625 8.941406 3.3406 9.1284
	$BladeCoords += HelperBladeProps slot4 7.09375 9.316406 8.90625 8.941406 5.4676 9.1284
	$BladeCoords += HelperBladeProps slot5 0.25 8.894531 2.0625 8.519531 3.3406 8.7065
	$BladeCoords += HelperBladeProps slot6 7.09375 8.894531 8.90625 8.519531 5.4676 8.7065
	$BladeCoords += HelperBladeProps slot7 0.25 8.476563 2.0625 8.101563 3.3406 8.2866
	$BladeCoords += HelperBladeProps slot8 7.09375 8.476563 8.90625 8.101563 5.4676 8.2866
	$BladeCoords += HelperBladeProps slot9 0.25 8.054688 2.0625 7.679688 3.3406 7.8647
	$BladeCoords += HelperBladeProps slot10 7.09375 8.054688 8.90625 7.679688 5.4676 7.8647
	$BladeCoords += HelperBladeProps slot11 7.09375 10.203125 8.90625 9.828125 4.392 10.016
	$BladeCoords += HelperBladeProps ps1 2.347656 7.125 3.410156 6.75 2.8799 7.4026
	$BladeCoords += HelperBladeProps ps2 3.480469 7.125 4.542969 6.75 4.0127 7.4026
	$BladeCoords += HelperBladeProps ps3 4.613281 7.125 5.675781 6.75 5.1455 7.4026
	$BladeCoords += HelperBladeProps ps4 5.746094 7.125 6.808594 6.75 6.2783 7.4026

	
	# Start Visio
	Write-Verbose "$VerbosePrefix Starting Visio"
	Start-Visio -Quiet
	
	# Load Stencil
	$StencilPath = "$PSScriptRoot\Resources\K-Series.vss"
	Write-Verbose "$VerbosePrefix Importing Stencil: $StencilPath"
	$LoadStencil = Import-VisioStencilFile $StencilPath
	
	###########################################################################################
	# Chassis and Hostname
	
	$ChassisModel = $DeviceObject.Hardware.PartNumber
	$ChassisName  = $DeviceObject.Name
	
	# Chassis Caption
	Write-Verbose "$VerbosePrefix Creating chassis caption"
	$Caption = Add-VisioRectangle 2.070313 10.75 7.070313 10.375 -Textbox -FontSize 18 -Text $ChassisModel
		
	# Chassis Stencil
	$Stencil = Select-VisioStencil $ChassisModel
	$Shape   = Add-VisioStencil $Stencil 4.57 8.7552

	# Hostname Caption
	$Caption = Add-VisioRectangle 2.070313 6.75 7.070313 6.375 -Textbox -FontSize 18 -Text $ChassisName
	
	###########################################################################################
	# Blades and OptionModules
	
	# Add Blades
	foreach ($s in $DeviceObject.Hardware.Slots) {
		$Slot = "slot" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		Write-Verbose "$VerbosePrefex $Slot"
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Model
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Model
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
	}
	
	###########################################################################################
	# Power Supplies
	
	foreach ($s in ($DeviceObject.Hardware.PowerSupplies | ? { $_.Status -ne "Not Installed"}) ) {
		$Slot = "ps" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Type
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Type
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
	}
	
	
	###########################################################################################
	# Clean Up and Save
	
	# Fit Page to Contents
	Set-VisioPageProperty -ResizeToFitContents
	
	# Save desired filetypes
	if ($VsdPath) { Save-VisioDocument $VsdPath	}
	if ($PngPath) { Export-VisioPage $PngPath -Resolution 300x300 }
	if ($SvgPath) { Export-VisioPage $SvgPath}
	
	# Display Visio if no file was saved, otherwise quit visio
	if ($VsdPath -or $PngPath -or $SvgPath) {
		Stop-Visio
	} else {
		$global:VisioShellInstance.App.Visible = $true
	}
}

###############################################################################
# New-ExVisioK6

function New-ExVisioK6 {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[ExtremeShell.Device]$DeviceObject,
		
		[Parameter(Mandatory=$False)]
		[string]$VsdPath,
		
		[Parameter(Mandatory=$False)]
		[string]$SvgPath,
		
		[Parameter(Mandatory=$False)]
		[string]$PngPath
	)
	
	# Used for Write-Verbose and Throw
	$VerbosePrefix = "New-ExVisioK6:"
	
	# Test for VisioShell
	$IsModuleAvailable = Get-Module -Listavailable -Name VisioShell
	if (!($IsModuleAvailable)) {
		Throw "$VerbosePrefex VisioShell Module not found."
	} else {
		$Import = Import-Module VisioShell
	}
	
	# Create objects for blade/psu placement
	$BladeCoords = @()
	$BladeCoords += HelperBladeProps slot1 0.25 9.8125 2.0625 9.4375 3.3564 9.6289
	$BladeCoords += HelperBladeProps slot2 7.09375 9.8125 8.90625 9.4375 5.4853 9.6289
	$BladeCoords += HelperBladeProps slot3 0.25 9.390625 2.0625 9.015625 3.3564 9.207
	$BladeCoords += HelperBladeProps slot4 7.09375 9.390625 8.90625 9.015625 5.4853 9.207
	$BladeCoords += HelperBladeProps slot5 0.25 8.96875 2.0625 8.59375 3.3564 8.7851
	$BladeCoords += HelperBladeProps slot6 7.09375 8.96875 8.90625 8.59375 5.4853 8.7851
	$BladeCoords += HelperBladeProps slot7 7.09375 10.28125 8.90625 9.90625 4.4195 10.0946
	$BladeCoords += HelperBladeProps ps1 2.359375 8.0625 3.421875 7.6875 2.8957 8.321
	$BladeCoords += HelperBladeProps ps2 3.5 8.0625 4.5625 7.6875 4.0304 8.321
	$BladeCoords += HelperBladeProps ps3 4.632813 8.0625 5.695313 7.6875 5.1632 8.321
	$BladeCoords += HelperBladeProps ps4 5.765625 8.0625 6.828125 7.6875 6.298 8.321
	
	# Start Visio
	Write-Verbose "$VerbosePrefix Starting Visio"
	Start-Visio -Quiet
	
	# Load Stencil
	$StencilPath = "$PSScriptRoot\Resources\K-Series.vss"
	Write-Verbose "$VerbosePrefix Importing Stencil: $StencilPath"
	$LoadStencil = Import-VisioStencilFile $StencilPath
	
	###########################################################################################
	# Chassis and Hostname
	
	$ChassisModel = $DeviceObject.Hardware.PartNumber
	$ChassisName  = $DeviceObject.Name
	
	# Chassis Caption
	Write-Verbose "$VerbosePrefix Creating chassis caption"
	$Caption = Add-VisioRectangle 2.09375 10.75 7.09375 10.375 -Textbox -FontSize 18 -Text $ChassisModel
		
	# Chassis Stencil
	$Stencil = Select-VisioStencil $ChassisModel
	$Shape   = Add-VisioStencil $Stencil 4.5863 9.2168

	# Hostname Caption
	$Caption = Add-VisioRectangle 2.09375 7.6875 7.09375 7.3125 -Textbox -FontSize 18 -Text $ChassisName
	
	###########################################################################################
	# Blades and OptionModules
	
	# Add Blades
	foreach ($s in $DeviceObject.Hardware.Slots) {
		$Slot = "slot" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		Write-Verbose "$VerbosePrefex $Slot"
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Model
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Model
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
	}
	
	###########################################################################################
	# Power Supplies
	
	foreach ($s in ($DeviceObject.Hardware.PowerSupplies | ? { $_.Status -ne "Not Installed"}) ) {
		$Slot = "ps" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Type
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Type
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
	}
	
	
	###########################################################################################
	# Clean Up and Save
	
	# Fit Page to Contents
	Set-VisioPageProperty -ResizeToFitContents
	
	# Save desired filetypes
	if ($VsdPath) { Save-VisioDocument $VsdPath	}
	if ($PngPath) { Export-VisioPage $PngPath -Resolution 300x300 }
	if ($SvgPath) { Export-VisioPage $SvgPath}
	
	# Display Visio if no file was saved, otherwise quit visio
	if ($VsdPath -or $PngPath -or $SvgPath) {
		Stop-Visio
	} else {
		$global:VisioShellInstance.App.Visible = $true
	}
}

###############################################################################
# New-ExVisioS4

function New-ExVisioS4 {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[ExtremeShell.Device]$DeviceObject,
		
		[Parameter(Mandatory=$False)]
		[string]$VsdPath,
		
		[Parameter(Mandatory=$False)]
		[string]$SvgPath,
		
		[Parameter(Mandatory=$False)]
		[string]$PngPath
	)
	
	# Used for Write-Verbose and Throw
	$VerbosePrefix = "New-ExVisioS4:"
	
	# Test for VisioShell
	$IsModuleAvailable = Get-Module -Listavailable -Name VisioShell
	if (!($IsModuleAvailable)) {
		Throw "$VerbosePrefex VisioShell Module not found."
	} else {
		$Import = Import-Module VisioShell
	}
	
	# Create objects for blade/psu placement
	$BladeCoords = @()
	$BladeCoords += HelperBladeProps slot1 5.3125 7.546875 8.75 7.1875 2.6349 7.534 4.2974 0.71
	$BladeCoords += HelperBladeProps slot2 5.3125 8.25 8.75 7.890625 2.6349 8.2455 4.2974 0.71
	$BladeCoords += HelperBladeProps slot3 5.3125 8.960938 8.75 8.601563 2.6339 8.9559 4.2974 0.71
	$BladeCoords += HelperBladeProps slot4 5.3125 9.671875 8.75 9.3125 2.6388 9.66 4.2974 0.71
	$BladeCoords += HelperBladeProps slot1_ul 5.3125 7.890625 7 7.546875 1.6356 7.7033 1.9423 0.213
	$BladeCoords += HelperBladeProps slot1_ur 7.0625 7.890625 8.75 7.546875 3.6298 7.7029 1.9423 0.213
	$BladeCoords += HelperBladeProps slot2_ul 5.3125 8.601563 7 8.25 1.6356 8.4148 1.9423 0.213
	$BladeCoords += HelperBladeProps slot2_ur 7.0625 8.601563 8.75 8.25 3.6298 8.4143 1.9423 0.213
	$BladeCoords += HelperBladeProps slot3_ul 5.3125 9.3125 7 8.960938 1.6346 9.1252 1.9423 0.213
	$BladeCoords += HelperBladeProps slot3_ur 7.0625 9.3125 8.75 8.960938 3.6288 9.1247 1.9423 0.213
	$BladeCoords += HelperBladeProps slot4_ul 5.3125 10.015625 7 9.671875 1.6395 9.8293 1.9423 0.213
	$BladeCoords += HelperBladeProps slot4_ur 7.0625 10.015625 8.75 9.671875 3.6337 9.8288 1.9423 0.213
	$BladeCoords += HelperBladeProps ps1 0.5625 6.5625 1.625 6.3125 1.0928 6.8546 1.2783 0.4913
	$BladeCoords += HelperBladeProps ps2 1.703125 6.5625 2.765625 6.3125 2.2314 6.8541 1.2783 0.4913
	$BladeCoords += HelperBladeProps ps3 2.84375 6.5625 3.90625 6.3125 3.3691 6.8522 1.2783 0.4913
	$BladeCoords += HelperBladeProps ps4 3.984375 6.5625 5.046875 6.3125 4.5137 6.8541 1.2783 0.4913
	
	# Start Visio
	Write-Verbose "$VerbosePrefix Starting Visio"
	Start-Visio -Quiet
	
	# Load Stencil
	$StencilPath = "$PSScriptRoot\Resources\S-Series.vss"
	Write-Verbose "$VerbosePrefix Importing Stencil: $StencilPath"
	$LoadStencil = Import-VisioStencilFile $StencilPath
	
	###########################################################################################
	# Chassis and Hostname
	
	$ChassisModel = $DeviceObject.Hardware.PartNumber
	$ChassisName  = $DeviceObject.Name
	
	# Chassis Caption
	Write-Verbose "$VerbosePrefix Creating chassis caption"
	$Caption = Add-VisioRectangle 0.265625 10.75 5.328125 10.375 -Textbox -FontSize 18 -Text $ChassisModel
		
	# Chassis Stencil
	$Stencil = Select-VisioStencil $ChassisModel
	$Shape   = Add-VisioStencil $Stencil 2.7984 8.4697

	# Hostname Caption
	$Caption = Add-VisioRectangle 0.265625 6.3125 5.328125 5.9375 -Textbox -FontSize 18 -Text $ChassisName
	
	###########################################################################################
	# Blades and OptionModules
	
	# Add Blades
	foreach ($s in $DeviceObject.Hardware.Slots) {
		$Slot = "slot" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		Write-Verbose "$VerbosePrefex $Slot"
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Model
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Model
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
		
		# OptionModules
		foreach ($o in $s.OptionModules) {
			$OptionSlot = $o.Location.Split()
			$OptionLoc  = $OptionSlot[0].SubString(0,1) + $OptionSlot[1].SubString(0,1)
			$OptionSlot = $Slot + '_' + $OptionLoc
			Write-Verbose "$VerbosePrefex $Slot"
			$b = $BladeCoords | ? { $_.slot -eq $OptionSlot }
			
			# Create Caption
			$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $o.Model
			
			# Change color for option Slots
			switch ($OptionLoc) {
				'ul' {
					Set-VisioShapeFont $Caption -ColorInHex C00000
				}
				'ur' {
					Set-VisioShapeFont $Caption -ColorInHex 0070C0
				}
			}
				
			# Select and Drop Blade Stencil
			$Stencil = Select-VisioStencil $o.Model
			$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
		}
	}
	
	###########################################################################################
	# Power Supplies
	
	foreach ($s in ($DeviceObject.Hardware.PowerSupplies | ? { $_.Status -ne "Not Installed"}) ) {
		$Slot = "ps" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		
		# Create Caption
		$Caption = Add-VisioRectangle $b.x1 $b.y1 $b.x2 $b.y2 -Textbox -FontSize 18 -Text $s.Type
			
		# Select and Drop Blade Stencil
		$Stencil = Select-VisioStencil $s.Type
		$Shape   = Add-VisioStencil $Stencil $b.pinx $b.piny
	}
	
	
	###########################################################################################
	# Clean Up and Save
	
	# Fit Page to Contents
	Set-VisioPageProperty -ResizeToFitContents
	
	# Save desired filetypes
	if ($VsdPath) { Save-VisioDocument $VsdPath	}
	if ($PngPath) { Export-VisioPage $PngPath -Resolution 300x300 }
	if ($SvgPath) { Export-VisioPage $SvgPath}
	
	# Display Visio if no file was saved, otherwise quit visio
	if ($VsdPath -or $PngPath -or $SvgPath) {
		Stop-Visio
	} else {
		$global:VisioShellInstance.App.Visible = $true
	}
}

###############################################################################
## Start Helper Functions
###############################################################################

###############################################################################
# HelperBladeProps

function HelperBladeProps {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Slot,

        [Parameter(Mandatory=$true,Position=1)]
        [double]$x1,

        [Parameter(Mandatory=$true,Position=2)]
        [double]$y1,

        [Parameter(Mandatory=$true,Position=3)]
        [double]$x2,

        [Parameter(Mandatory=$true,Position=4)]
        [double]$y2,

        [Parameter(Mandatory=$False,Position=5)]
        [double]$PinX,

        [Parameter(Mandatory=$False,Position=6)]
        [double]$PinY,

        [Parameter(Mandatory=$False,Position=7)]
        [double]$Width,

        [Parameter(Mandatory=$False,Position=8)]
        [double]$Height
    )

    $NewObject = "" | Select slot,x1,y1,x2,y2,pinx,piny,width,height
    $NewObject.slot   = $Slot
    $NewObject.x1     = $x1
    $NewObject.y1     = $y1
    $NewObject.x2     = $x2
    $NewObject.y2     = $y2
    $NewObject.PinX   = $PinX
    $NewObject.PinY   = $PinY
    $NewObject.Width  = $Width
    $NewObject.Height = $Height
    return $NewObject
}

###############################################################################
# HelperEvalRegex

function HelperEvalRegex {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$True,Position=0,ParameterSetName='RxString')]
		[String]$RegexString,
		
		[Parameter(Mandatory=$True,Position=0,ParameterSetName='Rx')]
		[regex]$Regex,
		
		[Parameter(Mandatory=$True,Position=1)]
		[string]$StringToEval,
		
		[Parameter(Mandatory=$False)]
		[string]$ReturnGroupName,
		
		[Parameter(Mandatory=$False)]
		[int]$ReturnGroupNumber,
		
		[Parameter(Mandatory=$False)]
		$VariableToUpdate,
		
		[Parameter(Mandatory=$False)]
		[string]$ObjectProperty,
		
		[Parameter(Mandatory=$False)]
		[string]$LoopName
	)
	
	$VerbosePrefix = "HelperEvalRegex: "
	
	if ($RegexString) {
		$Regex = [Regex] $RegexString
	}
	
	if ($ReturnGroupName) { $ReturnGroup = $ReturnGroupName }
	if ($ReturnGroupNumber) { $ReturnGroup = $ReturnGroupNumber }
	
	$Match = $Regex.Match($StringToEval)
	if ($Match.Success) {
		#Write-Verbose "$VerbosePrefix Matched: $($Match.Value)"
		if ($ReturnGroup) {
			#Write-Verbose "$VerbosePrefix ReturnGroup"
			switch ($ReturnGroup.Gettype().Name) {
				"Int32" {
					$ReturnValue = $Match.Groups[$ReturnGroup].Value.Trim()
				}
				"String" {
					$ReturnValue = $Match.Groups["$ReturnGroup"].Value.Trim()
				}
				default { Throw "ReturnGroup type invalid" }
			}
			if ($VariableToUpdate) {
				if ($VariableToUpdate.Value.$ObjectProperty) {
					#Property already set on Variable
					continue $LoopName
				} else {
					$VariableToUpdate.Value.$ObjectProperty = $ReturnValue
					Write-Verbose "$ObjectProperty`: $ReturnValue"
				}
				continue $LoopName
			} else {
				return $ReturnValue
			}
		} else {
			return $Match
		}
	} else {
		if ($ObjectToUpdate) {
			return
			# No Match
		} else {
			return $false
		}
	}
}

###############################################################################
## Export Cmdlets
###############################################################################

Export-ModuleMember *-*
