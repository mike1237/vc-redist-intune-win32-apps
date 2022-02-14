# requires:
#   Install-Module IntuneWin32App
#   Install-Module VcRedist

$intuneWinAppDir = "$PWD"
$SourceFolder = "$PWD\source"
$OutputFolder = "$PWD\output"
$warningPreference = "SilentlyContinue"

Write-Output "Downloading Source Packages..."
try{

    $redists = Get-VcList -Release 2022, 2013, 2012
    Save-VcRedist -VcList $redists -Path $SourceFolder
}
catch {
    Write-Output "Fail to download redists"
    exit 1
}



    ForEach ($redist in $redists) {
        write-output "====REDIST $($redist.Name) ===="

        # Package .intunewin
        $SetupFile = (Split-Path $redist.Download -Leaf)
        $Package = [System.IO.Path]::Combine($SourceFolder, $redist.Release, $redist.Version, $redist.Architecture)
        $Output = [System.IO.Path]::Combine($OutputFolder, $redist.Release, $redist.Version, $redist.Architecture)
        try {
            New-Item -Path $Output -ItemType "Directory"
        }
        Catch {
            Write-Output "Warning: Directory $Output already exists"
        }

        #Start-Process -FilePath "$intuneWinAppDir\IntuneWinAppUtil.exe" -ArgumentList "-c $Package -s $SetupFile -o $Output -q" -Wait -NoNewWindow
        
        Write-Output "Building App Package..."
        try {
            $Win32AppPackage = New-IntuneWin32AppPackage -SourceFolder $Package -SetupFile $SetupFile -Output $Output -Verbose
        }
        Catch {
            Write-Output "IntuneWin32App Packagine Fail"
            return 1
        }
        

        #Create app icon
        $AppIcon = New-IntuneWin32AppIcon -FilePath "$SourceFolder\Microsoft-VisualStudio.png"

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
            DisplayName              = "$Publisher Visual C++ Redistributable $($redist.Release) $($redist.Version) $($redist.Architecture)"
            Description              = "$Publisher $($redist.Name) $($redist.Version) $($redist.Architecture)."
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




