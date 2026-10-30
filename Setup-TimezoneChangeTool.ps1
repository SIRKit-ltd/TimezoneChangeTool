#requires -Version 7.2
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
try {
    if (-not $IsWindows) { throw 'This tool requires Windows.' }
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'Start setup using Setup-and-Launch.cmd.'
    }
    foreach ($name in 'Start-TimezoneChangeTool.ps1','TimezoneEngine.psm1','SeriesMigration.ps1','GraphRequestContract.ps1','config.json') {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name) -PathType Leaf)) {
            throw "Missing $name. Extract the entire package before running setup."
        }
    }

    Write-Host 'Execution-policy bypass applies to this session only. Organization policies still apply.'
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host 'Installing Microsoft.Graph.Authentication from PowerShell Gallery for the current user...'
        $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $repository) {
            Register-PSRepository -Default -ErrorAction Stop
            $repository = Get-PSRepository -Name PSGallery -ErrorAction Stop
        }
        if ($repository.SourceLocation.TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') {
            throw 'PSGallery points to an unexpected source. Ask IT to review the repository configuration.'
        }
        Install-Module -Name Microsoft.Graph.Authentication -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
    } else {
        Write-Host 'Microsoft.Graph.Authentication is already installed.'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Host 'Setup complete. Opening Timezone Change Tool...'
    & (Join-Path $PSScriptRoot 'Start-TimezoneChangeTool.ps1')
} catch {
    Write-Host ''
    Write-Host ('Setup or launch failed: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'For module installation failures, check internet/proxy access and your approved software policy.'
    Write-Host 'Organization-enforced script restrictions require IT approval; this launcher cannot override them.'
    exit 1
}
