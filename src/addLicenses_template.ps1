# Microsoft 365 PowerShell Script: Assign Licenses Based on Domain with Reporting and Device Code Authentication

# ==============================
# Configuration Variables
# ==============================

# Global Administrator Username
$adminUsername = "admin@mycompany.onmicrosoft.com"  # Replace with the email address of the Global Admin

# Domain to filter users (e.g., mycompany.com)
$domain = "mycompany.com"  # Replace with the desired domain

# SKU Identifiers to assign (e.g., SKU_LICENSE_IDENTIFIER_1, SKU_LICENSE_IDENTIFIER_2)
$skuIdentifiers = @("SKU_LICENSE_IDENTIFIER_1", "FLOW_SKU_LICENSE_IDENTIFIER_2FREE")  # Replace with the desired SKU Identifiers

# ==============================
# Script Starts Here
# ==============================

# Step 1: Log in to Microsoft 365 using Device Code Authentication
Write-Host "Logging in to Microsoft 365 using Device Code Authentication..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All"
    Write-Host "Successfully logged in to Microsoft 365." -ForegroundColor Green
} catch {
    Write-Host "Login failed. Please check your credentials or device code authentication process." -ForegroundColor Red
    return
}

# Step 2: List available SKUs with the number of licenses purchased and available
Write-Host "Fetching available licenses (SKUs)..." -ForegroundColor Cyan
try {
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

# Step 3: Get the SkuIds (GUIDs) of the selected SKUs
try {
    $skuIds = @()
    foreach ($skuIdentifier in $skuIdentifiers) {
        try {
            $sku = $null  # Explicitly initialize to null before comparison
            $sku = $availableSkus | Where-Object { $_.SkuPartNumber -eq $skuIdentifier }
            if ($null -eq $sku) {
                Write-Host "Invalid SKU Identifier entered: $skuIdentifier" -ForegroundColor Red
                Disconnect-MgGraph
                return
            }
            $skuIds += $sku.SkuId
            Write-Host "Selected SKU ID for "$skuIdentifier": $($sku.SkuId)" -ForegroundColor Green
        } catch {
            Write-Host "Error occurred while processing SKU Identifier: $skuIdentifier. Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "An unexpected error occurred. Error: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph  # Ensure cleanup happens on failure
    return
}

Write-Host ""
Write-Host "-------------------"
Write-Host ""

# Step 4: Filter users by domain (using Where-Object in PowerShell instead of $filter)
Write-Host "Fetching users with domain: $domain" -ForegroundColor Yellow
try {
    $domainUsers = Get-MgUser -All | Where-Object { $_.UserPrincipalName.EndsWith($domain) }
    if ($domainUsers.Count -eq 0) {
        Write-Host "No users found with the domain $domain!" -ForegroundColor Red
        Disconnect-MgGraph
        return
    } else {
        Write-Host "$($domainUsers.Count) users found with the domain $domain." -ForegroundColor Green
        Write-Host ""
        Write-Host "-------------------"
        Write-Host ""
    }
} catch {
    Write-Host "Error while fetching or filtering users: $_.Exception.Message" -ForegroundColor Red
    Disconnect-MgGraph
    return
}

# Step 5: Assign licenses and report
Write-Host "Starting license assignment for SKUs: $($skuIdentifiers -join ", ")" -ForegroundColor Yellow

$continueAll = $false  # Flag to track if the user chose "continue all"

foreach ($user in $domainUsers) {
    try {
        # Step 5.1: Get current licenses assigned to the user
        Write-Host "Fetching current licenses for user: $($user.UserPrincipalName)" -ForegroundColor Cyan

        # Get current licenses as SkuIds (GUIDs)
        $currentLicenses = (Get-MgUserLicenseDetail -UserId $user.Id).SkuId
        if ($null -eq $currentLicenses) {
            Write-Host "Current licenses are not available." -ForegroundColor Blue
        } else {
            Write-Host "Current licenses: $([string]::Join(', ', $currentLicenses))" -ForegroundColor Blue
        }

        # Map SkuPartNumber to SkuId for licenses to remove
        $existingSkuIds = $currentLicenses | Where-Object { $skuIds -notcontains $_ }

        # Prepare the licenses to add
        $licensesToAdd = $skuIds | ForEach-Object { @{SkuId = $_} }

        # Assign and remove licenses as needed
        if (-not $existingSkuIds -or $existingSkuIds.Count -eq 0) {
            Set-MgUserLicense -UserId $user.Id -AddLicenses $licensesToAdd -RemoveLicenses @()
        } else {
            Set-MgUserLicense -UserId $user.Id -AddLicenses $licensesToAdd -RemoveLicenses $existingSkuIds
        }

        Write-Host "Licenses assigned successfully with SKU IDs: $($skuIds -join ', ') to user: $($user.UserPrincipalName)" -ForegroundColor Green

        # Step 5.3: Fetch updated licenses
        Write-Host "Fetching updated licenses for user: $($user.UserPrincipalName)" -ForegroundColor Cyan
        $updatedLicenses = (Get-MgUserLicenseDetail -UserId $user.Id).SkuPartNumber
        Write-Host "Updated licenses: $([string]::Join(', ', $updatedLicenses))" -ForegroundColor Green

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
Disconnect-MgGraph
