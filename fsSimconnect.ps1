# Path where script is located
Write-Host "Running fsSimconnect: " -NoNewline
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

# If running on 64 bit OS make sure that 32 bit version of PowerShell is being used
if ((Get-WMIObject win32_OperatingSystem).OsArchitecture -match '64-bit') {
    if ([System.Diagnostics.Process]::GetCurrentProcess().Path -notmatch '\\syswow64\\') {
        throw 'Please run this script with 32-bit version of PowerShell'
    }
}

# Make sure that SP2 SIMCONNECT DLL is available
$ref = @($env:windir + '\assembly\GAC_32\Microsoft.FlightSimulator.SimConnect\10.0.61259.0__31bf3856ad364e35\Microsoft.FlightSimulator.SimConnect.dll')
if (!(Test-Path $ref)) {
    throw "Cannot find $ref. Make sure MFS is installed"
}

[System.Reflection.Assembly]::LoadFrom($ref) | Out-Null

try { $config = [xml](gc "$scriptPath\fsVar.xml") }
catch { throw 'Unable to find fsVar.xml in script folder, or fsVar.xml not in proper xml form' }

if (-not ('DataRequests' -as [type])) {
Add-Type -TypeDefinition @"
public enum DataRequests
    {
        Request1
    }
"@
}

if (-not ('Definitions' -as [type])) {
Add-Type -TypeDefinition @"
public enum Definitions
    {
        Struct1,
        Init
    }
"@
}

if (-not ('Struct1' -as [type])) {
    $type = 'using System.Runtime.InteropServices;'
    $type += '[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]'
    $type += 'public struct Struct1 {'
	
	foreach ($var in $config.fs.var) {
        $var.type = $var.type.ToLower()
        if ($var.type -eq 'string') { $type += "[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]" }
		$type += "public " + $var.type + ' ' + ($var.name -replace '[\s,:]','_') + ";"
    }
    $type += '}'
    Add-Type -TypeDefinition $type
}

if (-not ('EventId' -as [type])) {
    $type = 'public enum EventId {'
    foreach ($eventid in $config.fs.eventid) {
        $type += $eventid.name.ToUpper() + ","
    }
    $type += '}'
    Add-Type -TypeDefinition $type
}

if (-not ('Groups' -as [type])) {
Add-Type -TypeDefinition @"
   public enum Groups
   {
      group1
   }
"@
}

function transmit([EventId]$eventId, [int32]$param=0) {
    [uint32]$newParam = 0
    if ($param -lt 0) { $newParam = convert-IntToUint $param } else { $newParam = $param }
    $global:fs.TransmitClientEvent(0, $eventId, $newParam, [Groups]::group1, 
        [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_EVENT_FLAG]::GROUPID_IS_PRIORITY)
}

function convert-IntToUint([int32]$number) {
    $bytes = [bitconverter]::GetBytes($number)
    [bitconverter]::ToUInt32($bytes, 0)
}

try { $global:fs = New-Object Microsoft.FlightSimulator.SimConnect.SimConnect('PowerShell', 0, 1026, $null, 0) }
catch { throw 'Unable to connect to MFS. Please make sure MFS is started' }

[EventId].GetEnumValues() | % { 
    $global:fs.MapClientEventToSimEvent($_, $_.ToString())
}

$dataDefTypeMap = @{
    string = [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_DATATYPE]::STRING256
    int = [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_DATATYPE]::INT32
    long = [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_DATATYPE]::INT64
    float = [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_DATATYPE]::FLOAT32
    double = [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_DATATYPE]::FLOAT64
}

$global:simnames = @()
foreach ($var in $config.fs.var) {
    $var.type = $var.type.ToLower()
    if ($dataDefTypeMap.ContainsKey($var.type)) {
        $global:fs.AddToDataDefinition([Definitions]::Struct1, $var.name, $var.unit, $dataDefTypeMap[$var.type], 0, [Microsoft.FlightSimulator.SimConnect.SimConnect]::SIMCONNECT_UNUSED)
		$namefix = $var.name -replace '[\s,:]','_'
		$global:simnames += $namefix
	}
    else { throw 'Error: fsVar.xml: The Type for variable ' + $var.name + ' is not recognized' }
}

# Many thanks to Lee Holmes for the following three lines that enables calling generic SIMCONNECT function from PowerShell
# http://www.leeholmes.com/blog/2006/08/18/creating-generic-types-in-powershell/
$method = [Microsoft.FlightSimulator.SimConnect.SimConnect].GetMethod('RegisterDataDefineStruct')
$closedMethod = $method.MakeGenericMethod([Struct1])
$closedMethod.Invoke($global:fs, [Definitions]::Struct1)

$global:fsConnected = $true
$global:fsError = $false
Unregister-Event *
Register-ObjectEvent -InputObject $global:fs -EventName OnRecvException -Action { $global:fsError = $true } | out-null
Register-ObjectEvent -InputObject $global:fs -EventName OnRecvQuit -Action { $global:fsConnected = $false } | out-null
Register-ObjectEvent -InputObject $global:fs -EventName OnRecvSimobjectData -Action { try { $global:sim = $args.dwData[0] } catch {} } | out-null

$global:fs.RequestDataOnSimObject([DataRequests]::Request1, [Definitions]::Struct1, [Microsoft.FlightSimulator.SimConnect.SimConnect]::SIMCONNECT_OBJECT_ID_USER, [Microsoft.FlightSimulator.SimConnect.SIMCONNECT_PERIOD]::SIM_FRAME, 0, 0, 0, 0);

$timer = New-Object System.Timers.Timer
$timer.Interval = 100
$timer.AutoReset = $true
Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action { $global:fs.ReceiveMessage() } | out-null
$timer.Start()

if ($global:fsError) {
    $timer.Stop()
    $global:fs.Dispose()
    throw 'Error: fsVar.xml: There is at least one variable that SIMCONNECT does not recognize'
}

Write-Host "OK"

# Example transmit
# Set Autopilot altitude reference in feet
# transmit -eventId AP_ALT_VAR_SET_ENGLISH -param 4000

# Print simulation variables
# $global:sim

# Print altitude
# $global:sim.altitude
