# Microsoft Flight Simulator Logger (fsLogger)

## What it does
This script connects to a Microsoft Flight Simulator (FSX or newer), extracts flight variables (like a flight data recorder or "black box") and records it to a CSV file. The script can also automatically load KST2 for real-time viewing of the data while flying. It logs about 10-20 points per second, depending on the FPS, which is sufficient to resolve most manouvers and 

This can be used to analyze flight details, such as climb rate, G-forces, bank angles, etc., which is usefull as a student training to fly a real airplane or if you are a airplane/modding developer wanting to investigate flight behaviour in more detail.

## Requirements
* PowerShell v3 or higher
* FSX SP2 (or "newer")
* .Net 4.5 or higher
* (Otional) KST2 data visualization tool

## Installation
1. Place **fsLogger** files in any directory of your choice (where you have write access)
1. Install SimConnect (Microsoft API for communicating with the simulation)
   1. This is the same client used by for example Little Navmap (should be version 10.0.61259)
   1. Example of download location: https://www.littlenavmap.org/downloads/SimConnect/
1. (Optional) Download KST2 (https://kst-plot.kde.org/) and place the "Kst" folder in the same folder as the script so that the exe is located in "/kst/bin/KST2.exe"

## Running
Double click START.bat
This will open a PowerShell console in admin mode (you need admin rights) that connects to the simulator and it creates a Data.csv file which updates contineously as long as the simulation is running.

If you also have installed KST and the variable $startKST (in fsLogging.ps1) is set to "1", then KST will start and display all data in real time.

fsVar.xml contains all the variables that is extracted.
All variables that can be added or used can be found here: https://docs.microsoft.com/en-us/previous-versions/microsoft-esp/cc526981(v=msdn.10)

## References
The script is based on the PowerShell code in the FSX RESTful API project by pariljain (https://github.com/paruljain/fsx).

## Known issues
* SimConnect seems to only work for the first Flight Simulator installation (known bug in SimConnect)
* The script does not work properly if you change time of day mid flight, as it uses the time for detection of new datapoints.
   * Solution: Restart the script
* The script is quite simple, hence there is almost no error handling or other fancy features.

