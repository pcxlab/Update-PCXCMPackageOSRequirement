<#
.SYNOPSIS
    This script updates SCCM package programs to support a specified OS platform and generates a detailed report.

.DESCRIPTION
    The script reads a list of SCCM packages from a CSV file and checks if each package exists. If a package exists,
    it updates the supported OS platforms for all programs within the package. For each package and its programs,
    the script generates a detailed report including package and program details. If a package does not exist,
    it notes this in the report. If any specific package details are blank, they are replaced with values from the
    CSV file, if available. The report and a log file are saved with timestamps in their filenames.

.PARAMETER configFilePath
    Path to the configuration XML file containing the OS platform name.

.PARAMETER csvFilePath
    Path to the CSV file containing the list of package names and additional details to process.

.PARAMETER reportFilePath
    Path to the output CSV file where the report will be saved. A timestamp is appended to the filename.

.PARAMETER logFilePath
    Path to the log file where execution details will be saved. A timestamp is appended to the filename.

.NOTES
    The script requires the Configuration Manager module.

.EXAMPLE
    .\Automated_Package_OS_Requirement_Update.ps1

    This example runs the script with the default configuration, reading the package list from the current directory
    and generating a report and log file in the same directory.

#>

# Load required module
# Import-Module ConfigurationManager

Clear-Host

Set-Location $PSScriptRoot
##################################################################################################################
function Get-SCCMSiteCode {
    try {
        $siteCode = Get-WmiObject -Namespace "Root\SMS" -Class SMS_ProviderLocation -ComputerName "." | Select-Object -ExpandProperty SiteCode
        if ($siteCode -ne $null) {
            if ($siteCode -is [array]) {
                return $siteCode[0]
            } else {
                return $siteCode
            }
        } else {
            Write-Output "SCCM Site Code not found."
            return $null
        }
    } catch {
        Write-Error "Error retrieving SCCM Site Code: $_"
        return $null
    }
}

# Call the function and assign the result to a variable
$SiteCode = Get-SCCMSiteCode

# Check if the site code was retrieved successfully
if ($SiteCode -ne $null) {
    Write-Output "Assigned SCCM Site Code is: $SiteCode"
} else {
    Write-Output "Failed to retrieve SCCM Site Code."
    exit 0
}

# Get FQDN
function Get-SystemFQDN {
    try {
        $fqdn = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName
        if ($fqdn -ne $null) {
            return $fqdn
        } else {
            Write-Output "FQDN not found."
            return $null
        }
    } catch {
        Write-Error "Error retrieving FQDN: $_"
        return $null
    }
}

# Call the function and assign the result to a variable
$ProviderMachineName = Get-SystemFQDN

# Check if the FQDN was retrieved successfully
if ($ProviderMachineName -ne $null) {
    Write-Output "Assigned FQDN is: $ProviderMachineName"
} else {
    Write-Output "Failed to retrieve FQDN."
    exit 0
}

# Test Site Config
# $SiteCode = "PS1" # Site code 
# $ProviderMachineName = "CM01.corp.pcxlab.com" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors
# Do not change anything below this line
# Import the ConfigurationManager.psd1 module 
if ((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}
# Connect to the site's drive if it is not already present
if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}
# Set the current location to be the site code.
# Set-Location "$($SiteCode):\" @initParams
##################################################################################################################
function Set-CMProgramOSPlatform {
    param (
        [string]$packageName,
        [string]$programName,
        [string]$platform
    )

    $OSPlatform = Get-CMSupportedPlatform -Name $platform -Fast
    if (-not $OSPlatform) {
        throw "The specified platform '$platform' was not found."
    }

    Set-CMProgram -PackageName $packageName -ProgramName $programName -AddSupportedOperatingSystemPlatform $OSPlatform -StandardProgram
}

##################################################################################################################
# Configuration
$configFilePath = "$PSScriptRoot\Config.xml"
$csvFilePath = "$PSScriptRoot\Package_List.csv"
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$reportFilePath = "$PSScriptRoot\Automatic_Package_OS_Requirement_report_$timestamp.csv"
$logFilePath = "$PSScriptRoot\Automatic_Package_OS_Requirement_log_$timestamp.log"

# Start Transcript for logging
Start-Transcript -Path $logFilePath

# Read configuration from XML
[xml]$config = Get-Content -Path $configFilePath
$Platform = $config.Configuration.Platform

$siteCode = Get-SCCMSiteCode

# Set-Location $siteCode`:

Set-Location $PSScriptRoot

# Read the list of packages from CSV
$packages = Import-Csv -Path $csvFilePath

# Initialize report
$report = @()

# Function to log messages
function Write-Log {
    param (
        [string]$message
    )
    $logMessage = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $message"
    Write-Host $logMessage
}

# Define the target versions for Windows 11
$targetMaxVersion = "10.00.99999.9999"
$targetMinVersion = "10.00.22000.0"
$targetPlatform = "x64"

# Process each package
foreach ($package in $packages) {
    $packageName = $package."PackageName"
    Write-Log "Processing package: $packageName"

    Set-Location "$($SiteCode):\"

    $pkg = Get-CMPackage -Name $packageName -Fast -ErrorAction SilentlyContinue

    if ($pkg) {
        $programs = $pkg | Get-CMProgram
        if ($programs.Count -eq 0) {
            # If no programs found, log the package with "Not Found" in ProgramName column
            $status = "Not Found"
            $packageDetails = New-Object PSObject -Property @{
                PackageName           = $pkg.Name
                ProgramName           = $status
                Status                = ""
                Description           = $pkg.Description
                PackageID             = $pkg.PackageID
                Manufacturer          = $pkg.Manufacturer
                SourceSite            = $pkg.SourceSite
                PackageSize           = $pkg.PackageSize
                NoOfPrograms          = $programs.Count
                PackageSourcePath     = $pkg.PackageSourcePath
                PkgSourceFlag         = $pkg.PkgSourceFlag
                Priority              = $pkg.Priority
                ObjectPath            = $pkg.ObjectPath
                SourceDate            = $pkg.SourceDate
                TransformAnalysisDate = $pkg.TransformAnalysisDate
                SourceVersion         = $pkg.SourceVersion
                StoredPkgVersion      = $pkg.StoredPkgVersion
                LastRefreshTime       = $pkg.LastRefreshTime
            }

            # Replace blank values with CSV values if available
            foreach ($property in $packageDetails.PSObject.Properties) {
                if (-not $packageDetails.$($property.Name) -and $package.$($property.Name)) {
                    $packageDetails.$($property.Name) = $package.$($property.Name)
                }
            }

            $report += $packageDetails
        } else {
            foreach ($program in $programs) {
                # Retrieve the supported operating systems for the program
                $SupportedOperatingSystems = Get-CMProgram -PackageName $packageName -ProgramName $program.ProgramName | Select-Object -ExpandProperty SupportedOperatingSystems

                # Initialize a flag to track if Windows 11 is detected
                $windows11Detected = $false

                # Iterate through each supported operating system
                foreach ($os in $SupportedOperatingSystems) {
                    $supportedMaxVersion = $os.MaxVersion
                    $supportedMinVersion = $os.MinVersion
                    $supportedPlatform = $os.Platform

                    # Compare versions
                    if ($supportedMaxVersion -eq $targetMaxVersion -and
                        $supportedMinVersion -eq $targetMinVersion -and
                        $supportedPlatform -eq $targetPlatform) {
                        # Windows 11 version found
                        $windows11Detected = $true
                        break  # Exit the loop since we found a match
                    }
                }

                if ($windows11Detected) {
                    Write-Log "Program: $($program.ProgramName) in package: $packageName is already updated with Windows 11."
                    $status = "Already Updated with Windows 11"
                } else {
                    try {
                        Write-Log "Updating program: $($program.ProgramName) in package: $packageName"
                        Set-CMProgramOSPlatform -packageName $packageName -programName $program.ProgramName -platform $Platform
                        $status = "Updated"
                    } catch {
                        Write-Log "Failed to update program: $($program.ProgramName) in package: $packageName. Error: $_"
                        $status = "Update Failed"
                    }
                }

                $packageDetails = New-Object PSObject -Property @{
                    PackageName           = $pkg.Name
                    ProgramName           = $program.ProgramName
                    Status                = $status
                    Description           = $pkg.Description
                    PackageID             = $pkg.PackageID
                    Manufacturer          = $pkg.Manufacturer
                    SourceSite            = $pkg.SourceSite
                    PackageSize           = $pkg.PackageSize
                    NoOfPrograms          = $programs.Count
                    PackageSourcePath     = $pkg.PackageSourcePath
                    PkgSourceFlag         = $pkg.PkgSourceFlag
                    Priority              = $pkg.Priority
                    ObjectPath            = $pkg.ObjectPath
                    SourceDate            = $pkg.SourceDate
                    TransformAnalysisDate = $pkg.TransformAnalysisDate
                    SourceVersion         = $pkg.SourceVersion
                    StoredPkgVersion      = $pkg.StoredPkgVersion
                    LastRefreshTime       = $pkg.LastRefreshTime
                }

                # Replace blank values with CSV values if available
                foreach ($property in $packageDetails.PSObject.Properties) {
                    if (-not $packageDetails.$($property.Name) -and $package.$($property.Name)) {
                        $packageDetails.$($property.Name) = $package.$($property.Name)
                    }
                }

                $report += $packageDetails
            }
        }
    } else {
        Write-Log "Package not found: $packageName"
        $status = "Not Found"
        $packageDetails = New-Object PSObject -Property @{
            PackageName           = $packageName
            ProgramName           = ""
            Status                = $status
            Description           = if ($package.Description) { $package.Description } else { "" }
            PackageID             = if ($package.PackageID) { $package.PackageID } else { "" }
            Manufacturer          = if ($package.Manufacturer) { $package.Manufacturer } else { "" }
            SourceSite            = if ($package.SourceSite) { $package.SourceSite } else { "" }
            PackageSize           = if ($package.PackageSize) { $package.PackageSize } else { "" }
            NoOfPrograms          = if ($package.NoOfPrograms) { $package.NoOfPrograms } else { "" }
            PackageSourcePath     = if ($package.PackageSourcePath) { $package.PackageSourcePath } else { "" }
            PkgSourceFlag         = if ($package.PkgSourceFlag) { $package.PkgSourceFlag } else { "" }
            Priority              = if ($package.Priority) { $package.Priority } else { "" }
            ObjectPath            = if ($package.ObjectPath) { $package.ObjectPath } else { "" }
            SourceDate            = if ($package.SourceDate) { $package.SourceDate } else { "" }
            TransformAnalysisDate = if ($package.TransformAnalysisDate) { $package.TransformAnalysisDate } else { "" }
            SourceVersion         = if ($package.SourceVersion) { $package.SourceVersion } else { "" }
            StoredPkgVersion      = if ($package.StoredPkgVersion) { $package.StoredPkgVersion } else { "" }
            LastRefreshTime       = if ($package.LastRefreshTime) { $package.LastRefreshTime } else { "" }
        }

        $report += $packageDetails
    }
}

# Export the report to CSV in the specified order
$report | Select-Object PackageName, ProgramName, Status, Description, PackageID, Manufacturer, SourceSite, PackageSize, NoOfPrograms, PackageSourcePath, PkgSourceFlag, Priority, ObjectPath, SourceDate, TransformAnalysisDate, SourceVersion, StoredPkgVersion, LastRefreshTime | Export-Csv -Path $reportFilePath -NoTypeInformation

# Stop Transcript for logging
Stop-Transcript

Write-Log "Report generation completed. Report file: $reportFilePath"
Write-Log "Log file: $logFilePath"
