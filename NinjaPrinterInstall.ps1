#Modification of the script from Graham Pocta posted here: https://ninjarmm.zendesk.com/hc/en-us/community/posts/7102933603213--Full-printer-management-script

##############################################################################################
#  2022-06-15 Graham Pocta                                                                   #
#   Per-client documentation has to exist for printers to be installed via this script       #
#   Download print driver, create TCP/IP printer port, install printer object                #
#                                                                                            #
#  Requirements:                                                                             #
#   -Documentation template set up for Printers and Printers To Delete                       #
#   -Tested on Windows 10/11 x64                                                             #
#   -If installing multiple printers using the same driver, make sure they use the same      #
#    version number                                                                          #
##############################################################################################

$webClient = New-Object net.webclient

#Gets host's current gateway addresses
$Gateways = Get-NetIPConfiguration | Foreach IPv4DefaultGateway | Select -ExpandProperty NextHop

#Returns list of printers in org
$Printers = Ninja-Property-Docs-Names "Printers"

$PrinterCounter = 0

foreach ($Printer in $Printers){
	$PrinterCount++
	Write-Host "#############"
	Write-Host "# Printer $PrinterCount #"
	Write-Host "#############`n`n"
	
	#split the results string to get just the name of the printer in documentation
	$pos = $Printer.IndexOf("=")
	$printerName = $Printer.Substring($pos+1)
	Write-Host "Name: $printerName"
	$printerIP = Ninja-Property-Docs-Get "Printers" "$printerName" "ipAddress"
	Write-Host "IP: $printerIP"
	$driveFile = Ninja-Property-Docs-Get "Printers" "$printerName" "driverFile"
	Write-Host "Driver Filename: $driverFile"
	$driverURL = Ninja-Property-Docs-Get "Printers" "$printerName" "driverUrl"
	Write-Host "Driver URL: $driverURL"
	$driverExeName = Ninja-Property-Docs-Get "Printers" "$printerName" "driverExeName"
	if($driverExeName){Write-Host "Driver Switches: $driverExeName"}
	$newDriverVersion = Ninja-Property-Docs-Get "Printers" "$printerName" "driverVersion"
	Write-Host "Driver version: $newDriverVersion"
	$newDriverName = Ninja-Property-Docs-Get "Printers" "$printerName" "driverName"
	Write-Host "Driver name: $newDriverName"
	$specifyInfFile = Ninja-Property-Docs-Get "Printers" "$printerName" "specifyInfFile"
	Write-Host "INF file: $specifyInfFile"
	$filterToNetworksWithGateway = Ninja-Property-Docs-Get "Printers" "$printerName" "filterToNetworksWithGateway"
	if($filterToNetworksWithGateway){Write-Host "Networks: $filterToNetworksWithGateway"}
	$filterToHostnames = Ninja-Property-Docs-Get "Printers" "$printerName" "filterToHostnames"
	if($filterToHostnames){Write-Host "Hostnames: $filterToHostnames"}
	$duplexMode = Ninja-Property-Docs-Get "Printers" "$printerName" "duplexMode"

	Write-Host "Duplex setting: $duplexMode"
	$color = Ninja-Property-Docs-Get "Printers" "$printerName" "color"
	Write-Host "Color setting: $color"
	$collate = Ninja-Property-Docs-Get "Printers" "$printerName" "collate"
	Write-Host "Collate setting: $collate"

	$tempFolder = "C:\Temp"
	Write-Host "Temp folder: $tempFolder"
	$downloadFile = "$tempFolder\$driverFile"
	Write-Host "Temp driver download: $downloadFile"
	$hostname = Get-Content Env:COMPUTERNAME
	Write-Host "Current hostname: $hostname"
	
	#Remove leading zeros in version name, because the published version number sometimes includes "01" but then shows up as just "1" in the system
	$newDriverVersion=($newDriverVersion.split(".") | ForEach({ $_ -replace '^0+(?=[^.])' }) ) -join "."

	#Find out if printer already exists
	$printerExists = Get-Printer -Name "$printerName" -ErrorAction SilentlyContinue
	$printerDriverExists = Get-PrinterDriver -Name "$newDriverName" -ErrorAction SilentlyContinue
	
	#Print driver with the new driver name already exists 
	if($printerDriverExists){
		$oldDriverName = $newDriverName
	}

	#First we assume that this host is not on the list for hostnames or networks
	$thisPrinterShouldBeInstalled = $false

	if($filterToNetworksWithGateway){
		#We now iterate through all gateways as defined in the documentation
		Write-Host "Network filtering enabled. Checking that this host has a matching gateway."
		Foreach ($filterGateway in $filterToNetworksWithGateway){
			 Foreach ($Gateway in $Gateways){
					if($Gateway -eq $filterGateway){
						 #We found a match between the documentation-defined subnets and a current subnet on this host! Installation should proceed
						 Write-Host "Subnet match for $Gateway! Printer $printerName will be installed."
						 $thisPrinterShouldBeInstalled = $true
					}
			 }
		}
	}

	if($filterToHostnames){
		#We now iterate through all hostnames as defined in the documentation

		Foreach ($filterToHostname in $filterToHostnames){
			 if($hostname -eq $filterToHostname){
						 #We found a match between the documentation-defined hostname and this computer's hostname, the installation should proceed
						 Write-Host "Hostname match! Printer $printerName will be installed."
						 $thisPrinterShouldBeInstalled = $true
					}
			 }
		}
	
	if(-not $filterToNetworksWithGateway -and -not $filterToHostnames){
		#Since there is no filtering, this printer will be deployed to all hosts where the script runs
		$thisPrinterShouldBeInstalled = $true
	}

	if(-not $thisPrinterShouldBeInstalled){
		#Skipping installation, ensure printer is not installed on this host
		Write-Host "Skipping printer $printerName - it should not be installed to this host's current subnet or hostname."
		
		#Checking if printer already exists. If so, remove printer
		if($printerExists){
				Write-Host "Printer $printerName exists, but is filtered out by hostname or subject. Removing..."
				Remove-Printer -Name "$printerName" -ErrorAction SilentlyContinue
		}
		#Just some console text cleanliness
		Write-Host "Printer $PrinterCount complete. `n`n"
		
		#Continue on to next printer in documentation
		continue
	}
	
	#We now know that the printer SHOULD exist on this host, so we want to check 
	#if the driver name or version is different, in which case we will download,
	#install and apply the new driver. First, we assume that we won't do any of that
	$installNewDriver = $false
	$upgradeDriver = $false


	
	#if print driver exists, collect the version number
	if(-not $printerDriverExists)
	{
		Write-Host "Driver $newDriverName not found. New driver will be installed."
		$installNewDriver = $true
	}else{
		Write-Host "Driver of same name as new driver already exists, getting version of old driver..."
		$oldDriverVersion = Get-PrinterDriver -Name "$oldDriverName" | Select @{
			n="DriverVersion";e={
				$ver = $_.DriverVersion
				$rev = $ver -band 0xffff
				$build = ($ver -shr 16) -band 0xffff
				$minor = ($ver -shr 32) -band 0xffff
				$major = ($ver -shr 48) -band 0xffff
				"$major.$minor.$build.$rev"
			}
		} | Select -ExpandProperty DriverVersion
		Write-Host "Old driver version: $oldDriverVersion"
		if($newDriverVersion -ne $oldDriverVersion){
			Write-Host "Old driver version doesn't match new driver version"
			$upgradeDriver = $true
		}
		else{
			
		}
	}
			

	
	#If driver doesn't exist or installed version is different, we need to download/extract the driver files
	if($installNewDriver -or $upgradeDriver){
		#Clean up existing print driver download
		New-Item -ItemType Directory $tempFolder -ErrorAction SilentlyContinue
		Remove-Item $downloadFile -Force -ErrorAction SilentlyContinue
		Remove-Item $tempFolder\newDriver -Force -Recurse -ErrorAction SilentlyContinue
		New-Item -ItemType Directory $tempFolder\newDriver -ErrorAction SilentlyContinue
	
		#Download driver file to $tempFolder
		$webClient.DownloadFile($driverURL, $downloadFile)
		#Invoke-WebRequest -Uri "$driverURL" -OutFile $downloadFile
				
		Write-Host "Extracting $downloadFile to $tempFolder\newDriver"
				
		#Extract zip file containing drivers 
		#& 'C:\Program Files\7-Zip\7z.exe' x $downloadFile "-o$tempFolder\newDriver" -y
		tar -xf $downloadFile -C $tempFolder\newDriver
		####################################################################
		# Extract zip file containing drivers - this method is supposed		#
		# to work but is commented out because I ran into "file in use"		#
		# errors caused by our antivirus (Trend Micro)										 #
		####################################################################
		#Add-Type -Assembly "System.IO.Compression.Filesystem"
		#[System.IO.Compression.ZipFile]::ExtractToDirectory('$downloadFile','$tempFolder\newDriver')
		
		if($driverExeName) {
			Start-Process (Get-ChildItem -Recurse -Filter $driverExeName | Select-Object -First 1).fullname -ArgumentList "/VERYSILENT /DIR=$tempFolder\newDriver" 
		}

		Write-Host "Driver archive extracted"
	}
	
	if($installNewDriver -or $upgradeDriver){
				Write-Host "Driver $newDriverName $newDriverVersion is not installed. Installing..."
		
				#For clean installs
				if(-not $upgradeDriver){
				if(-not $specifyInfFile){
						#Install all drivers in expanded folder
						Write-Host "Installing all drivers in driver folder to system..."
						Get-ChildItem "$tempFolder\newDriver" -Recurse -Filter "*.inf" | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
				}
		
				#Install only specified INF file
				if($specifyInfFile){
						PNPUtil.exe /add-driver (Get-ChildItem "$tempFolder\newDriver" -Recurse -Filter "$specifyInfFile" | Select-Object -First 1).FullName /install
				}
				#Add requested driver name to local print server
				$newDriverInfPath = Get-WindowsDriver -Online -All | Where-Object {$_.ClassName -eq "Printer" -and $_.Version -eq "$newDriverVersion"} | select -ExpandProperty OriginalFileName
				Write-Host "Installing print driver: $newDriverName"
				Write-Host "New driver inf path: $newDriverInfPath"
				Add-PrinterDriver -Name "$newDriverName" -InfPath "$newDriverInfPath"
		}
	}

	#Check if printer port already exists. If not, add it
	$portExists = Get-Printerport -Name $printerIP -ErrorAction SilentlyContinue
	if (-not $portExists) {
		Write-Host "Adding printer port: $printerIP"
		Add-PrinterPort -name "$printerIP" -PrinterHostAddress "$printerIP"
	}

	#Check if printer driver already exists on local print server. If not, add it
	$printDriverExists = Get-PrinterDriver -name "$newDriverName" -ErrorAction SilentlyContinue
	if ($printDriverExists) {
		#Check if printer exists. If not, add it. If so, update printer port
		$printerExists = Get-Printer -name "$printerName" -ErrorAction SilentlyContinue
		if (-not $printerExists){
			Write-Host "Installing printer... this can take a few minutes"
			Add-Printer -Name $printerName -PortName "$printerIP" -DriverName "$newDriverName"
			}
		else{
			Write-Host "Printer already exists. Will update port and driver`n"
			
			Write-Host "Setting port to $printerIP"
			Set-Printer -Name $printerName -PortName "$printerIP"
			
			Write-Host "Setting driver to $newDriverName"
			Set-Printer -Name $printerName -DriverName "$newDriverName"
			}
				
		}
	else{
		Write-Warning "Error: Printer driver not installed, please check version number and URL to driver is correct"
	}
	
	#Set printer defaults
	switch ($duplexMode) {
		c9f09750-6b84-4a5b-b0b1-7ef9d6187bd1 {$duplexMode = "OneSided"}
		316e6a9a-22da-4d36-bf2b-858b911c5c45 {$duplexMode = "TwoSidedLongEdge"}
		79d617c5-d809-4ad6-b29d-2509afe8f297 {$duplexMode = "TwoSidedShortEdge"}
	}
	Set-PrintConfiguration $printerName -DuplexingMode $duplexMode -Color $color -Collate $collate

	#Just some console text cleanliness
	Write-Host "Printer $PrinterCount complete. `n`n"
}

#Clean up driver download directory regardless of if the driver was downloaded
Write-Host "Cleaning up temp files..."
Remove-Item $downloadFile -Force -ErrorAction SilentlyContinue
Remove-Item $tempFolder\newDriver -Force -Recurse -ErrorAction SilentlyContinue

Write-Host "#######################"
Write-Host "# Old printer cleanup #"
Write-Host "#######################`n`n"

#Find printers to delete from documentation and remove them
Write-Host "Deleting unwanted printers..."
$printersToDelete = Ninja-Property-Docs-Names "Printers To Delete"
foreach ($Printer in $printersToDelete){
	#split the results string to get just the name of the printer in documentation
	$pos = $Printer.IndexOf("=")
	$printerName = $Printer.Substring($pos+1)
		
	if(Get-Printer -Name "$printerName" -ErrorAction SilentlyContinue){
		#remove all print jobs
		Write-Host "Removing all printer jobs for printer: $printerName"
		Get-PrintJob -PrinterName "$printerName" | Remove-PrintJob
		
		#delete printer
		Write-Host "Deleting: $printerName..."
		Remove-Printer -Name "$printerName" -ErrorAction SilentlyContinue
	}
}

#Find print servers to delete from documentation and remove all associated printers
Write-Host "Deleting printers from unwanted print servers..."

#Build batch file to run on start-up
$printServers = Ninja-Property-Docs-Names "Print Servers To Delete"
if($printServers){
	#Path to script has to be location accessible to all users
	$deletePrintServersScriptPath = "C:\Users\Public\deletePrintServers.ps1"
	$ScriptContent = $null
	foreach ($printServer in $printServers){
		#split the results string to get just the name of the print server in documentation
		$pos = $printServer.IndexOf("=")
		$printServerName = $printServer.Substring($pos+1)
		$ScriptContent = $ScriptContent + "Remove-Printer -name \\$printServerName\*`n"
	}

	#Create (and erase existing) script file
	New-Item -Path $deletePrintServersScriptPath -ItemType "File" -Force | Set-Content -Value $ScriptContent
	
	#Schedule script to run at login
	schtasks /create /f /ru builtin\users /sc onlogon /tn "Remove printers from old print servers" /tr "powershell.exe -WindowStyle Hidden -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File $deletePrintServersScriptPath"
	$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
	Set-ScheduledTask -TaskName "Remove printers from old print servers" -Settings $settings
}else{
	#Delete scheduled task if no printers from any print server exist
	if(Get-ScheduledTask -TaskName "Remove printers from old print servers" -ErrorAction SilentlyContinue){
		Unregister-ScheduledTask -TaskName "Remove printers from old print servers" -Confirm:$False
	}
}

#Clean up unused printer drivers
$Drivers = Get-PrinterDriver
$Printers = Get-Printer

ForEach($Driver in $Drivers)
{
	$PrintersUsingDriver = ($Printers | Where {$_.DriverName -eq $Driver.Name} | Measure).Count
	Write-Host "Driver: $($Driver.Name) has $PrintersUsingDriver printers using it."

	if($PrintersUsingDriver -eq 0){
		try{
			Remove-PrinterDriver -Name $Driver.Name -ErrorAction Stop
			Write-Host "$($Driver.Name) has been removed."
		}
		catch{
			Write-Host "Failed to remove driver: $($Driver.Name)`n"
		}
	}
}

#Turn off automatic printer discovery
$path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private'
$key = try {
		Get-Item -Path $path -ErrorAction Stop
}
catch {
		New-Item -Path $path -Force
		Write-Host "Creating registry key $path..."
}
New-ItemProperty -Path $path -Name 'AutoSetup' -Value 0 -PropertyType Dword -Force
