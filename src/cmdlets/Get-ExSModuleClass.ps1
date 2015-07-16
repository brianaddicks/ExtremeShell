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

