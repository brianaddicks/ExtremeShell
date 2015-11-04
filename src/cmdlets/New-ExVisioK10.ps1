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