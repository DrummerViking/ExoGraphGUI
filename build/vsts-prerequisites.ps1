param (
    [string]
    $Repository = 'PSGallery'
)

$modules = @("Pester", "PSFramework", "BurntToast", "PSModuleDevelopment", "PSScriptAnalyzer", "Microsoft.Graph.Authentication", "Microsoft.Graph.Mail", "Microsoft.Graph.Users.Actions")

# Automatically add missing dependencies
$data = Import-PowerShellDataFile -Path "$PSScriptRoot\..\ExoGraphGUI\ExoGraphGUI.psd1"
foreach ($dependency in $data.RequiredModules) {
    if ($dependency -is [string]) {
        if ($modules -contains $dependency) { continue }
        $modules += $dependency
    }
    else {
        if ($modules -contains $dependency.ModuleName) { continue }
        $modules += $dependency.ModuleName
    }
}

foreach ($module in $modules) {
    Write-Host "Installing $module" -ForegroundColor Cyan
    Install-Module $module -Force -SkipPublisherCheck -Repository $Repository
    Import-Module $module -Force -PassThru
}