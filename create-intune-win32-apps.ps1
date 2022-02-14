# requires:
#   Install-Module IntuneWin32App
#   Install-Module VcRedist

$SupportedVcVers = 2022, 2013, 2012 # Release Years
$UnsupportedVcVer = "9.0.30729.6161" # Release Version
$SourceFolder = "$PWD\source"
$OutputFolder = "$PWD\output"

#todo
#document, detect if authenticated to azure, add build-only flag, add download-only flag, auto rewrite output file
function main {
    # Supported Pkgs
    $redists = Get-VcList -Release $SupportedVcVers
    Get-SourcePkgs $redists
    Add-IntuneWinPkgs $redists

    # Unsupported Pkgs
    # Info: https://vcredist.com/get-vclist/#returning-supported-redistributables
    $redists = Get-VcList -Export -All | Where-Object { $_.Version -eq $UnsupportedVcVer }
    Get-SourcePkgs $redists
    Add-IntuneWinPkgs $redists

    exit    
}

function Get-SourcePkgs {
    param($redists)
    Write-Output "Download Source Packages..."
    Try {
        Save-VcRedist -VcList $redists -Path $SourceFolder
    }
    Catch {
        Write-Output "Failed to download Source Packages!!"
        Return 1
    }
    Return 0
}

function Add-IntuneWinPkgs {
    param($redists)
    ForEach ($redist in $redists) {
        $SetupFile = (Split-Path $redist.Download -Leaf)
        $Package = [System.IO.Path]::Combine($SourceFolder, $redist.Release, $redist.Version, $redist.Architecture)
        $Output = [System.IO.Path]::Combine($OutputFolder, $redist.Release, $redist.Version, $redist.Architecture)
        if (!(Test-Path -Path $Output)) {
            New-Item -Path $Output -ItemType "Directory"
        }
        try {
            $Win32AppPackage = New-IntuneWin32AppPackage -SourceFolder $Package -SetupFile $SetupFile -Output $Output -Verbose
        }
        Catch {
            Write-Output "IntuneWin32App Packagine Fail"
            return 1
        }
        

        #Create app icon
        $AppIcon = New-IntuneWin32AppIcon -FilePath "$PWD\Microsoft-VisualStudio.png"

        # Enable 'Associated with a 32-bit app on 64-bit clients' as required
        $KeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        $params = @{
            Existence            = $true
            KeyPath              = "$KeyPath\$($redist.ProductCode)"
            Check32BitOn64System = If ($redist.UninstallKey -eq "32") { $True } Else { $False }
            DetectionType        = "exists"
        }
        $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry @params

        # Application requirement rules
        Switch ($redist.Architecture) {
            "x86" {
                $PackageArchitecture = "All"
            }
            "x64" {
                $PackageArchitecture = "x64"
            }
        }
        $params = @{
            Architecture                    = $PackageArchitecture
            MinimumSupportedOperatingSystem = "1607"
        }
        $RequirementRule = New-IntuneWin32AppRequirementRule @params

        # Add Intune App
        $Publisher = "Microsoft"
        $params = @{
            FilePath                 = $Win32AppPackage.Path
            DisplayName              = "$Publisher Visual C++ $($redist.Release) $($redist.Architecture) $($redist.Version)  Redistributable"
            Description              = "$Publisher $($redist.Name) $($redist.Architecture) $($redist.Version)."
            Publisher                = $Publisher
            InformationURL           = $redist.URL
            PrivacyURL               = "https://go.microsoft.com/fwlink/?LinkId=521839"
            CompanyPortalFeaturedApp = $false
            InstallExperience        = "system"
            RestartBehavior          = "suppress"
            DetectionRule            = $DetectionRule
            RequirementRule          = $RequirementRule
            InstallCommandLine       = ".\$SetupFile $($redist.SilentInstall)"
            UninstallCommandLine     = $redist.SilentUninstall
            Icon                     = $AppIcon
            Verbose                  = $true
        }
        Try
        {
            Add-IntuneWin32App @params
        } 
        Catch
        {
            Write-Output "Failed to upload package to Intune - Is Graph Authenticated?"
            Return 1
        }

    }
    Return 0
}
# Entrypoint
main

# sources:
# https://vcredist.com/new-vcintuneapplication/#package-the-redistributables
# https://msendpointmgr.com/2020/02/05/install-visual-c-redistributables-for-microsoft-intune-managed-devices/
# https://github.com/MSEndpointMgr/IntuneWin32App#full-example-of-packaging-and-creating-a-win32-app
# https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool


