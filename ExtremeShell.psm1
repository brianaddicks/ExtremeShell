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
## Start Helper Functions
###############################################################################

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
