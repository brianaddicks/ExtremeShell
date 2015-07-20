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
	$LoadStencil = Import-VisioStencilFile "$PSScriptRoot\Resources\S-Series.vss"
	
	###########################################################################################
	# Chassis and Hostname
	
	$ChassisModel = $DeviceObject.Hardware.PartNumber
	$ChassisName  = $DeviceObject.Name
	
	# Chassis Caption
	Write-Verbose "$VerbosePrefix Creating chassis caption"
	$Caption = Add-VisioRectangle 0.265625 10.75 5.328125 10.375 -Textbox -FontSize 18 -Text $ChassisName
		
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