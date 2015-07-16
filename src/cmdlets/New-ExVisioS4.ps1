function New-ExVisioS4 {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Takes the output of "show hardware" from an Enterasys switch and translates it into a Powershell Object.
	#>

	Param (
		[Parameter(Mandatory=$True,Position=0)]
		[ExtremeShell.Device]$DeviceObject,
		
		[Parameter(Mandatory=$True,Position=0)]
		[string]$StencilPath
	)
	
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
	$VisioApp = New-Object -ComObject Visio.Application
	$VisioApp.visible = $true
	
	# Create new Document
	$Documents = $VisioApp.Documents
	$Document = $Documents.Add("")
	
	# Select Active Pages
	$Pages = $VisioApp.ActiveDocument.Pages
	$Page  = $Pages.Item(1)
	
	# Load Stencil
	$LoadStencil = $VisioApp.Documents.Add($StencilPath)
	
	# Add Blades
	foreach ($s in $DeviceObject.Hardware.Slots) {
		$Slot = "slot" + $s.Number
		$b = $BladeCoords | ? { $_.slot -eq $Slot }
		Write-Verbose $Slot
		
		# Create Caption
		$Caption = $Page.DrawRectangle($b.x1, $b.y1, $b.x2, $b.y2)
		$Caption.TextStyle = "Normal"
	    $Caption.LineStyle = "Text Only"
	    $Caption.FillStyle = "Text Only"
		$Caption.Text      = $s.Model
		$Caption.CellsSRC(3,0,7).Formula = "18 pt"
			
		# Select and Drop Blade Stencil
		$Stencil = $LoadStencil.Masters.Item($s.Model)
		$Shape   = $Page.Drop($Stencil, $b.pinx, $b.piny)
		
		# OptionModules
		foreach ($o in $s.OptionModules) {
			$OptionSlot = $o.Location.Split()
			$OptionSlot = $OptionSlot[0].SubString(0,1) + $OptionSlot[1].SubString(0,1)
			$OptionSlot = $Slot + '_' + $OptionSlot
			Write-Verbose $OptionSlot
			$b = $BladeCoords | ? { $_.slot -eq $OptionSlot }
			
			# Create Caption
			$Caption = $Page.DrawRectangle($b.x1, $b.y1, $b.x2, $b.y2)
			$Caption.TextStyle = "Normal"
		    $Caption.LineStyle = "Text Only"
		    $Caption.FillStyle = "Text Only"
			$Caption.Text      = $o.Model
			$Caption.CellsSRC(3,0,7).Formula = "18 pt"
				
			# Select and Drop Blade Stencil
			$Stencil = $LoadStencil.Masters.Item($o.Model)
			$Shape   = $Page.Drop($Stencil, $b.pinx, $b.piny)
		}
	}
}