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

        [Parameter(Mandatory=$true,Position=5)]
        [double]$PinX,

        [Parameter(Mandatory=$true,Position=6)]
        [double]$PinY,

        [Parameter(Mandatory=$true,Position=7)]
        [double]$Width,

        [Parameter(Mandatory=$true,Position=8)]
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