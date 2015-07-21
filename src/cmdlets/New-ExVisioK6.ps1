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