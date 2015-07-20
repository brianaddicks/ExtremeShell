[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False,Position=0)]
	[switch]$PushToStrap
)

$ScriptPath = Split-Path $($MyInvocation.MyCommand).Path
$ModuleName = Split-Path $ScriptPath -Leaf

$SourceDirectory = "src"
$SourcePath      = $ScriptPath + "\" + $SourceDirectory
$CmdletPath      = $SourcePath + "\" + "cmdlets"
$HelperPath      = $SourcePath + "\" + "helpers"
$CsPath          = $SourcePath + "\" + "cs"
$OutputFile      = $ScriptPath + "\" + "$ModuleName.psm1"
$ManifestFile    = $ScriptPath + "\" + "$ModuleName.psd1"
$DllFile         = $ScriptPath + "\" + "$ModuleName.dll"
$CsOutputFile    = $ScriptPath + "\" + "$ModuleName.cs"

###############################################################################
# Create Manifest
$ManifestParams = @{ Path = $ManifestFile
                     ModuleVersion = '1.0'
                     RequiredAssemblies = @("$ModuleName.dll",'System.Web')
                     Author             = 'Brian Addicks'
                     RootModule         = "$ModuleName.psm1"
                     PowerShellVersion  = '4.0' 
                     RequiredModules    = @('ipv4math')}

New-ModuleManifest @ManifestParams

###############################################################################
# 

$CmdletHeader = @'
###############################################################################
## Start Powershell Cmdlets
###############################################################################


'@

$HelperFunctionHeader = @'
###############################################################################
## Start Helper Functions
###############################################################################


'@

$Footer = @'
###############################################################################
## Export Cmdlets
###############################################################################

Export-ModuleMember *-*
'@

$FunctionHeader = @'
###############################################################################
# 
'@

###############################################################################
# Start Output

$CsOutput  = ""

###############################################################################
# Add C-Sharp

$AssemblyRx       = [regex] '^using\ .+?;'
$NameSpaceStartRx = [regex] "namespace $ModuleName {"
$NameSpaceStopRx  = [regex] '^}$'

$Assemblies    = @()
$CSharpContent = @()

$c = 0
foreach ($f in $(ls $CsPath)) {
    foreach ($l in (gc $f.FullName)) {
        $AssemblyMatch       = $AssemblyRx.Match($l)
        $NameSpaceStartMatch = $NameSpaceStartRx.Match($l)
        $NameSpaceStopMatch  = $NameSpaceStopRx.Match($l)

        if ($AssemblyMatch.Success) {
            $Assemblies += $l
            continue
        }

        if ($NameSpaceStartMatch.Success) {
            $AddContent = $true
            continue
        }

        if ($NameSpaceStopMatch.Success) {
            $AddContent = $false
            continue
        }

        if ($AddContent) {
            $CSharpContent += $l
        }
    }
}

#$Assemblies | Select -Unique | sort -Descending

$CSharpOutput  = $Assemblies | Select -Unique | sort -Descending
$CSharpOutput += "namespace $ModuleName {"
$CSharpOutput += $CSharpContent
$CSharpOutput += '}'

$CsOutput += [string]::join("`n",$CSharpOutput)
$CsOutput | Out-File $CsOutputFile -Force


Add-Type -ReferencedAssemblies @(
	([System.Reflection.Assembly]::LoadWithPartialName("System.Xml")).Location,
	([System.Reflection.Assembly]::LoadWithPartialName("System.Web")).Location,
	([System.Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq")).Location
	) -OutputAssembly $DllFile -OutputType Library -TypeDefinition $CsOutput

###############################################################################
# Add Cmdlets

$Output = $CmdletHeader

foreach ($l in $(ls $CmdletPath)) {
    $Contents  = gc $l.FullName
    $Output   += $FunctionHeader
    $Output   += $l.BaseName
    $Output   += "`r`n`r`n"
    $Output   += [string]::join("`n",$Contents)
    $Output   += "`r`n`r`n"
}


###############################################################################
# Add Helpers

$Output += $HelperFunctionHeader

foreach ($l in $(ls $HelperPath)) {
    $Contents  = gc $l.FullName
    $Output   += $FunctionHeader
    $Output   += $l.BaseName
    $Output   += "`r`n`r`n"
    $Output   += [string]::join("`n",$Contents)
    $Output   += "`r`n`r`n"
}

###############################################################################
# Add Footer

$Output += $Footer

###############################################################################
# Output File

$Output | Out-File $OutputFile -Force

###############################################################################
# Copy to Strap

if ($PushToStrap) {
    # Create Temporary folder for zipping
    $TempZipFolder = 'newzip'
    $TempZipFullPath = "$($env:temp)\$TempZipFolder"
    $CreateFolder = New-Item -Path $env:temp -Name $TempZipFolder -ItemType Directory
    
    # Select Files for Zipping
    $FilesToZip = ls "$PSScriptRoot\$ModuleName*" -Exclude *.zip
    $Copy       = Copy-Item $FilesToZip -Destination $TempZipFullPath
    if (Test-Path "$PSScriptRoot\Resources") {
        $CopyResources = Copy-Item "$PSScriptRoot\Resources" -Recurse -Destination $TempZipFullPath
    } 
    
    # Zip them Up
    $ZipFilePath = "$PSScriptRoot\$ModuleName.zip"
    $Delete      = Remove-Item $ZipFilePath
    
    Add-Type -Assembly System.IO.Compression.FileSystem
    $CompressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
    [System.IO.Compression.ZipFile]::CreateFromDirectory($TempZipFullPath,$ZipFilePath, $CompressionLevel, $false)
    
    # Copy to Strap
    $StageFolder = "\\vmware-host\Shared Folders\Dropbox\strap\stages\$ModuleName\"
        
    try {
        $CheckForFolder = ls $StageFolder -ErrorAction Stop
    } catch {
        $MakeFolder = New-Item -Path $StageFolder -ItemType Directory
    }
    $Copy = Copy-Item "$PSScriptRoot\$ModuleName.zip" $StageFolder -Force
    $Remove = Remove-Item $TempZipFullPath -Recurse
    
        # Create stage_init.ps1
    $StageInitContents  = @()
    $StageInitContents += "#DESCRIPTION $ModuleName`r`n`r`n"
    $StageInitContents += "Push-Location -path $ModuleName`r`n`r`n"
    $StageInitContents += "if (!(Test-Path .\$ModuleName.psd1)) {`r`n" 
    
    $StageInitContents += @"
 	#not downloaded and extracted; do it
	
	`$shell_app=new-object -com shell.application
	`$filename = "$ModuleName.zip"
	`$zip_file = `$shell_app.namespace((Get-Location).Path + "\`$filename")
	`$(`$shell_app.namespace((Get-Location).Path)).Copyhere(`$zip_file.items())
	Remove-Item `$filename
	Remove-Variable shell_app,filename,zip_file
}
Pop-Location
"@
    
    $StageInitContents | Out-File "$StageFolder\stage_init.ps1"
}