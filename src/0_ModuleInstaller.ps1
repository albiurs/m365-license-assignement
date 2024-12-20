# ==============================
# Comment here
# ==============================

# Check if any Microsoft.Graph related modules are loaded
$loadedModules = Get-Module | Where-Object { $_.Name -like "Microsoft.Graph*" }
if ($loadedModules) {
    Write-Host "Removing conflicting Microsoft.Graph modules..." -ForegroundColor Yellow
    # Unload all Microsoft Graph related modules
    $loadedModules | ForEach-Object { Remove-Module -Name $_.Name -Force }
}

# Ensure that the Microsoft.Graph module is imported
Write-Host "Importing the Microsoft.Graph module..." -ForegroundColor Yellow
Import-Module Microsoft.Graph -Force
