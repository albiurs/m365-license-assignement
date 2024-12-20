# Microsoft 365 PowerShell Script: Assign Licenses Based on Domain with Reporting and Device Code Authentication

# ==============================
# Configuration Variables
# ==============================

# Global Administrator Username
$adminUsername = "digitagadm@schulezizersch.onmicrosoft.com"  # Replace with the email address of the Global Admin

# Domain to filter users (e.g., mycompany.com)
$domain = "schueler.zizers.ch"  # Replace with the desired domain

# SKU Identifier to assign (e.g., Office 365 A1 for Students)
$skuIdentifier = "STANDARDWOFFPACK_STUDENT"  # Replace with the desired SKU Identifier

# ==============================
# Script Starts Here
# ==============================

# Step 1: Log in to Microsoft 365 using Device Code Authentication
Write-Host "Logging in to Microsoft 365 using Device Code Authentication..." -ForegroundColor Cyan
try {
    # Attempt to connect using Microsoft Graph
    Connect-MgGraph -Scopes "User.ReadWrite.All"
    Write-Host "Successfully logged in to Microsoft 365." -ForegroundColor Green
} catch {
    Write-Host "Login failed. Please check your credentials or device code authentication process." -ForegroundColor Red
    return
}

# ==============================
# Fix Module Conflicts
# ==============================

# # Check if any Microsoft.Graph related modules are loaded
# $loadedModules = Get-Module | Where-Object { $_.Name -like "Microsoft.Graph*" }
# if ($loadedModules) {
#     Write-Host "Removing conflicting Microsoft.Graph modules..." -ForegroundColor Yellow
#     # Unload all Microsoft Graph related modules
#     $loadedModules | ForEach-Object { Remove-Module -Name $_.Name -Force }
# }

# # Ensure that the Microsoft.Graph module is imported
# Write-Host "Importing the Microsoft.Graph module..." -ForegroundColor Yellow
# Import-Module Microsoft.Graph -Force

# Step 2: List available SKUs with the number of licenses purchased and available
Write-Host "Fetching available licenses (SKUs)..." -ForegroundColor Cyan
try {
    # Fetch subscribed SKUs, including purchased and available units
    $availableSkus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, 
        @{Name="Purchased Licenses";Expression={$_.PrepaidUnits.Enabled}}, 
        @{Name="Available Licenses";Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}} 

    Write-Host "Available SKUs:" -ForegroundColor Yellow
    $availableSkus | Format-Table -Property SkuPartNumber, "Purchased Licenses", "Available Licenses" -AutoSize
} catch {
    Write-Host "Error fetching SKUs: $_.Exception.Message" -ForegroundColor Red
    Disconnect-MgGraph
    return
}

# Step 3: Get the SkuId (GUID) of the selected SKU
$sku = $availableSkus | Where-Object { $_.SkuPartNumber -eq $skuIdentifier }

if ($sku -eq $null) {
    Write-Host "Invalid SKU Identifier entered." -ForegroundColor Red
    Disconnect-MgGraph
    return
}

$skuId = $sku.SkuId
Write-Host "Selected SKU ID: $skuId" -ForegroundColor Green

# Step 4: Filter users by domain (using Where-Object in PowerShell instead of $filter)
Write-Host "Fetching users with domain: $domain" -ForegroundColor Yellow
try {
    # Fetch all users, and filter them locally by checking the userPrincipalName
    $domainUsers = Get-MgUser -All | Where-Object { $_.UserPrincipalName.EndsWith($domain) }

    if ($domainUsers.Count -eq 0) {
        Write-Host "No users found with the domain $domain!" -ForegroundColor Red
        Disconnect-MgGraph
        return
    } else {
        Write-Host "$($domainUsers.Count) users found with the domain $domain." -ForegroundColor Green
    }
} catch {
    Write-Host "Error while fetching or filtering users: $_.Exception.Message" -ForegroundColor Red
    Disconnect-MgGraph
    return
}

# Step 5: Assign licenses and report
Write-Host "Starting license assignment for SKU: $skuIdentifier" -ForegroundColor Yellow

$continueAll = $false  # Flag to track if the user chose "continue all"

foreach ($user in $domainUsers) {
    try {
        # Step 5.1: Get current licenses assigned to the user
        Write-Host "Fetching current licenses for user: $($user.UserPrincipalName)" -ForegroundColor Cyan
        $currentLicenses = (Get-MgUserLicenseDetail -UserId $user.Id).SkuPartNumber
        Write-Host "Current licenses: $([string]::Join(', ', $currentLicenses))" -ForegroundColor Blue

        # Step 5.2: Assign the new license
        Write-Host "Assigning new license ($skuIdentifier) to user: $($user.UserPrincipalName)" -ForegroundColor Yellow

        # Assign license with an empty array for RemoveLicenses, ensuring no licenses are removed
        Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $skuId} -RemoveLicenses @()
        Write-Host "License assigned successfully to user: $($user.UserPrincipalName)" -ForegroundColor Green

        # Step 5.3: Fetch updated licenses
        Write-Host "Fetching updated licenses for user: $($user.UserPrincipalName)" -ForegroundColor Cyan
        $updatedLicenses = (Get-MgUserLicenseDetail -UserId $user.Id).SkuPartNumber
        Write-Host "Updated licenses: $([string]::Join(', ', $updatedLicenses))" -ForegroundColor Green

        # Print an empty line for separation
        Write-Host ""
        Write-Host "-------------------"
        Write-Host ""

        # Step 5.4: Prompt for continuation if not in "continue all" mode
        if (-not $continueAll) {
            $continue = Read-Host "Do you want to continue to the next user? (y/n/a for continue all)"
            if ($continue -eq "n") {
                Write-Host "Stopping the script as requested." -ForegroundColor Yellow
                break
            } elseif ($continue -eq "a") {
                $continueAll = $true
            }
        }

    } catch {
        Write-Host "Error while assigning license to user: $($user.UserPrincipalName)" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

# Completion
Write-Host "License assignment complete. Script finished." -ForegroundColor Cyan

# Log out of Microsoft 365
Disconnect-MgGraph
