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