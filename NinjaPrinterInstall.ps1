$tempFolder = "$env:TEMP\printer"
$log = "$tempFolder\printer.log"
Start-Transcript -Path $log
#Modification of the script from Graham Pocta posted here: https://ninjarmm.zendesk.com/hc/en-us/community/posts/7102933603213--Full-printer-management-script

##############################################################################################
#  2022-06-15 Graham Pocta                                                                   #
#   Per-client documentation has to exist for printers to be installed via this script       #
#   Download print driver, create TCP/IP printer port, install printer object                #
#                                                                                            #
#  Requirements:                                                                             #
#   -Documentation template set up for Printers and Printers To Delete                       #
#   -Tested on Windows 10 x64                                            					 #
#   -If installing multiple printers using the same driver, make sure they use the same      #
#    version number                                                                          #
##############################################################################################

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
	$driverFile = Ninja-Property-Docs-Get "Printers" "$printerName" "driverFilename"
	Write-Host "Driver Filename: $driverFile"
	$driverURL = Ninja-Property-Docs-Get "Printers" "$printerName" "driverUrl"
	Write-Host "Driver URL: $driverURL"
	$newDriverVersion = Ninja-Property-Docs-Get "Printers" "$printerName" "driverVersion"
	Write-Host "Driver version: $newDriverVersion"
	$newDriverName = Ninja-Property-Docs-Get "Printers" "$printerName" "driverName"
	Write-Host "Driver name: $newDriverName"
	$specifyInfFile = Ninja-Property-Docs-Get "Printers" "$printerName" "specifyInfFile"
	Write-Host "INF file: $specifyInfFile"
	$duplexMode = Ninja-Property-Docs-Get "Printers" "$printerName" "duplexMode"
	Write-Host "Duplex setting: $duplexMode"
	$color = Ninja-Property-Docs-Get "Printers" "$printerName" "color"
	Write-Host "Color setting: $color"
	$collate = Ninja-Property-Docs-Get "Printers" "$printerName" "collate"
	Write-Host "Collate setting: $collate"
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
	
	#We now know that the printer SHOULD exist on this host, so we want to check 
	#if the driver name or version is different, in which case we will download,
	#install and apply the new driver. First, we assume that we won't do any of that
	$installNewDriver = $false
	$upgradeDriver = $false
	
	#if print driver exists, collect the version number
	if(-not $printerDriverExists) {
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
		}else{
			Write-Host "Old driver version matches new driver version"
			$upgradeDriver = $false
		}
	}
	
	#If driver doesn't exist or installed version is different, we need to download/extract the driver files
	if($installNewDriver -or $upgradeDriver){
		Write-Host "Clean up existing print driver download"
		New-Item -ItemType Directory $tempFolder -ErrorAction SilentlyContinue
		Remove-Item $downloadFile -Force -ErrorAction SilentlyContinue
		Remove-Item $tempFolder\newDriver -Force -Recurse -ErrorAction SilentlyContinue
		New-Item -ItemType Directory $tempFolder\newDriver -ErrorAction SilentlyContinue
	
		Write-Host "Download driver file to $tempFolder"
		(New-Object Net.WebClient).DownloadFile($driverURL, $downloadFile)
				
		Write-Host "Extracting $downloadFile to $tempFolder\newDriver"
		switch ((Get-ChildItem $downloadFile).Extension) {
			.zip {Write-Host "Extracting from zip"; tar -xf $downloadFile -C $tempFolder\newDriver}
			.exe {Write-Host "Extracting from exe"; Start-Process $downloadFile -ArgumentList "/VERYSILENT /DIR=$tempFolder\newDriver" }
			Default {Write-Warning "Error extracting drive, quiting"; exit}
		}

		Write-Host "Driver archive extracted"
	}
	
	if($installNewDriver -or $upgradeDriver){
		Write-Host "Driver $newDriverName $newDriverVersion is not installed. Installing..."
		#For clean installs
		if(-not $upgradeDriver){
			#Install all drivers in expanded folder
			Write-Host "Installing all drivers in driver folder to system..."
			Get-ChildItem "$tempFolder\newDriver" -Recurse -Filter "*.inf" | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }
		}
		
		#Add requested driver name to local print server
		$newDriverInfPath = Get-WindowsDriver -Online -All | Where-Object {$_.ClassName -eq "Printer" -and $_.Version -eq "$newDriverVersion"} | select -ExpandProperty OriginalFileName
		Write-Host "Installing print driver: $newDriverName"
		Write-Host "New driver inf path: $newDriverInfPath"
		Add-PrinterDriver -Name "$newDriverName" -InfPath "$newDriverInfPath"
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
		}else{
			Write-Host "Printer already exists. Will update port and driver`n"
			
			Write-Host "Setting port to $printerIP"
			Set-Printer -Name $printerName -PortName "$printerIP"
			
			Write-Host "Setting driver to $newDriverName"
			Set-Printer -Name $printerName -DriverName "$newDriverName"
			}
				
		}else{
		Write-Warning "Error: Printer driver not installed, please check version number and URL to driver is correct"
	}
	
	Write-Host "Setting print configuration"
	switch ($duplexMode) {
		c9f09750-6b84-4a5b-b0b1-7ef9d6187bd1 {$duplexMode = "OneSided"}
		316e6a9a-22da-4d36-bf2b-858b911c5c45 {$duplexMode = "TwoSidedLongEdge"}
		79d617c5-d809-4ad6-b29d-2509afe8f297 {$duplexMode = "TwoSidedShortEdge"}
	}

	switch ($color) {
		0 {[bool]$color=$false}
		1 {[bool]$color=$true}
		Default {Write-Warning "Error setting color config, exiting"; exit}
	}
	
	switch ($collate) {
		0 {[bool]$collate=$false}
		1 {[bool]$collate=$true}
		Default {Write-Warning "Error setting collate config, exiting"; exit}
	}

	Set-PrintConfiguration $printerName -DuplexingMode $duplexMode -Color $color -Collate $collate

	#Just some console text cleanliness
	Write-Host "Printer $PrinterCount complete. `n`n"
}

#Clean up driver download directory regardless of if the driver was downloaded
Write-Host "Cleaning up temp files..."
Remove-Item $downloadFile -Force -ErrorAction SilentlyContinue
Remove-Item $tempFolder\newDriver -Force -Recurse -ErrorAction SilentlyContinue

#Clean up unused printer drivers
$Drivers = Get-PrinterDriver | where {$_.name -ne "Microsoft enhanced Point and Print compatibility driver"}
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

Stop-Transcript
