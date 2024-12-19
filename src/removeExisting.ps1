# Step 6.2: Assign the new license
Write-Host "Assigning new license ($skuPartNumber) to user: $($user.UserPrincipalName)" -ForegroundColor Yellow

# Fetch current licenses to prevent duplication (optional)
$currentLicenses = (Get-MgUserLicenseDetail -UserId $user.Id).SkuId
$existingSkuIds = $currentLicenses | Where-Object { $_ -ne $skuId }

# Assign license, ensure we also provide the 'removeLicenses' parameter
Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $skuId} -RemoveLicenses $existingSkuIds
Write-Host "License assigned successfully to user: $($user.UserPrincipalName)" -ForegroundColor Green